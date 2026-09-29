from __future__ import annotations

from io import BytesIO
import re

import pandas as pd
import requests
import streamlit as st


API_BASE = "http://127.0.0.1:8530"
MONTH_RE = re.compile(r"^(20\\d{2})/(0?[1-9]|1[0-2])$")
RAW_COLUMNS = [
    "previous_month",
    "debit",
    "credit",
    "current_month",
    "previous_year",
    "diff",
    "ratio",
]


def month_label(key: int) -> str:
    year, month = divmod(key, 100)
    return f"{year}年{month}月"


def previous_key(key: int) -> int:
    year, month = divmod(key, 100)
    return (year - 1) * 100 + 12 if month == 1 else year * 100 + month - 1


def read_trend_csv(content: bytes) -> tuple[pd.DataFrame, dict[int, str]]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            frame = pd.read_csv(BytesIO(content), encoding=encoding)
            break
        except UnicodeError as exc:
            last_error = exc
    else:
        raise ValueError(f"CSVの文字コードを判別できません: {last_error}")

    frame.columns = [str(column).strip() for column in frame.columns]
    code_column = next((column for column in frame.columns if "科目コード" in column), None)
    name_column = next((column for column in frame.columns if "科目名" in column), None)
    if code_column is None or name_column is None:
        raise ValueError("勘定科目コード・勘定科目名の列が見つかりません。")

    month_columns = {}
    for column in frame.columns:
        match = MONTH_RE.fullmatch(column)
        if match:
            month_columns[int(match.group(1)) * 100 + int(match.group(2))] = column
    if not month_columns:
        raise ValueError("YYYY/MM形式の月列が見つかりません。")

    frame = frame.rename(columns={code_column: "account_code", name_column: "account_name"})
    frame["account_code"] = frame["account_code"].map(
        lambda value: "" if pd.isna(value) else str(value).strip().removesuffix(".0")
    )
    frame["account_name"] = frame["account_name"].fillna("").astype(str).str.strip()
    frame = frame[(frame["account_code"] != "") & (frame["account_name"] != "")].copy()

    duplicates = frame.loc[frame["account_code"].duplicated(keep=False), "account_code"].unique()
    if len(duplicates):
        raise ValueError("勘定科目コードが重複しています: " + ", ".join(duplicates[:10]))

    for column in month_columns.values():
        values = frame[column].astype(str).str.replace(",", "", regex=False).str.strip()
        frame[column] = pd.to_numeric(values, errors="coerce")
    return frame, month_columns


def fetch_balances(client_name: str, key: int) -> dict[str, dict]:
    year, month = divmod(key, 100)
    response = requests.get(
        f"{API_BASE}/balances",
        params={
            "client_name": client_name,
            "period_year": year,
            "period_month": month,
            "limit": 5000,
        },
        timeout=20,
    )
    response.raise_for_status()
    return {str(row["account_code"]): row for row in response.json()}


def build_period_values(
    source: pd.DataFrame,
    month_columns: dict[int, str],
    closing_month: int,
    prior_periods: dict[int, dict[str, dict]],
) -> dict[int, dict[str, float]]:
    """Convert P&L month movements into fiscal YTD values for the database."""
    fiscal_start = closing_month % 12 + 1
    period_values: dict[int, dict[str, float]] = {}
    source_keys = set(month_columns)
    for key in sorted(month_columns):
        previous = previous_key(key)
        if key % 100 == fiscal_start:
            base: dict[str, float] = {}
        elif previous in period_values:
            base = period_values[previous]
        else:
            base = {
                code: float(row.get("current_month") or 0)
                for code, row in prior_periods.get(previous, {}).items()
            }

        values: dict[str, float] = {}
        month_column = month_columns[key]
        for _, account in source.iterrows():
            code = str(account["account_code"])
            amount = account[month_column]
            amount = float(amount) if pd.notna(amount) else 0.0
            values[code] = (
                base.get(code, 0.0) + amount
                if code[:1] in "456789"
                else amount
            )
        period_values[key] = values
    return period_values


def build_month_frame(
    source: pd.DataFrame,
    key: int,
    month_columns: dict[int, str],
    period_values: dict[int, dict[str, float]],
    existing: dict[str, dict],
    prior_periods: dict[int, dict[str, dict]],
) -> pd.DataFrame:
    previous = previous_key(key)
    previous_rows = prior_periods.get(previous, {})
    previous_values = period_values.get(previous, {})
    month_column = month_columns[key]
    output = []

    for _, account in source.iterrows():
        code = str(account["account_code"])
        old_row = existing.get(code, {})
        amount = account[month_column]
        monthly_amount = float(amount) if pd.notna(amount) else 0.0
        current = period_values[key].get(code, 0.0)
        prior_value = previous_values.get(
            code,
            float(previous_rows.get(code, {}).get("current_month") or 0),
        )

        if code[:1] in "456789":
            if code.startswith("4"):
                debit, credit = max(-monthly_amount, 0.0), max(monthly_amount, 0.0)
            else:
                debit, credit = max(monthly_amount, 0.0), max(-monthly_amount, 0.0)
        else:
            debit = float(old_row.get("debit") or 0)
            credit = float(old_row.get("credit") or 0)

        previous_year = float(old_row.get("previous_year") or 0)
        output.append(
            {
                "account_code": code,
                "account_name": str(account["account_name"]),
                "previous_month": prior_value,
                "debit": debit,
                "credit": credit,
                "current_month": current,
                "previous_year": previous_year,
                "diff": current - previous_year,
                "ratio": current / previous_year * 100 if previous_year else 0.0,
            }
        )
    return pd.DataFrame(output, columns=["account_code", "account_name", *RAW_COLUMNS])


st.title("勘定科目残高推移表から再登録")
st.caption("CSVの月別値を照合し、選んだ関与先・年月だけを財務API経由で置き換えます。")
st.info(
    "この画面は既存の財務分析アプリ内で動作します。利用端末をTailscaleに接続し、"
    "いつもの財務分析アプリからこのページを開いてください。"
)

try:
    clients_response = requests.get(f"{API_BASE}/clients", timeout=10)
    clients_response.raise_for_status()
    clients = clients_response.json()
except requests.RequestException as exc:
    st.error(f"財務API（{API_BASE}）に接続できません: {exc}")
    st.stop()

uploaded = st.file_uploader("勘定科目残高推移表CSV", type=["csv"])
if uploaded is None:
    st.info("CSVを選ぶと、含まれる年月と既存データの差を表示します。")
    st.stop()

try:
    source, month_columns = read_trend_csv(uploaded.getvalue())
except Exception as exc:
    st.error(f"CSVを読み込めません: {exc}")
    st.stop()

client_options = {
    f"{client['client_id']}｜{client['client_name']}": client
    for client in clients
}
selected_client_label = st.selectbox(
    "再登録する関与先（CSVに会社名がないため選択してください）",
    list(client_options),
    index=None,
    placeholder="関与先を選択してください",
)
if selected_client_label is None:
    st.stop()

selected_client = client_options[selected_client_label]
client_name = selected_client["client_name"]
closing_month = int(selected_client.get("fiscal_year_start_month") or 12)
keys = sorted(month_columns)
selected_keys = st.multiselect(
    "再登録する年月",
    keys,
    default=keys,
    format_func=month_label,
    help="選択した月だけを置き換えます。CSVにない科目の既存行は削除されます。",
)
if not selected_keys:
    st.warning("再登録する月を1つ以上選んでください。")
    st.stop()

try:
    prior_keys = {previous_key(key) for key in keys if previous_key(key) not in month_columns}
    prior_periods = {key: fetch_balances(client_name, key) for key in prior_keys}
    period_values = build_period_values(source, month_columns, closing_month, prior_periods)
    snapshots: dict[int, tuple[dict[str, dict], pd.DataFrame]] = {}
    preview = []

    for key in selected_keys:
        existing = fetch_balances(client_name, key)
        month_frame = build_month_frame(
            source, key, month_columns, period_values, existing, prior_periods
        )
        snapshots[key] = (existing, month_frame)
        sales = month_frame.loc[month_frame["account_code"] == "4000", "current_month"]
        old_sales = existing.get("4000", {}).get("current_month")
        new_sales = float(sales.iloc[0]) if not sales.empty else 0.0
        preview.append(
            {
                "年月": month_label(key),
                "CSV科目数": len(month_frame),
                "DB登録科目数": len(existing),
                "売上高累計（CSV計算）": new_sales,
                "売上高累計（DB）": float(old_sales) if old_sales is not None else None,
                "差額": new_sales - float(old_sales) if old_sales is not None else None,
            }
        )
except Exception as exc:
    st.error(f"DBとの照合に失敗しました: {exc}")
    st.stop()

st.subheader("置換内容の確認")
st.dataframe(pd.DataFrame(preview), hide_index=True, use_container_width=True)
new_periods = [month_label(key) for key, (rows, _) in snapshots.items() if not rows]
if new_periods:
    st.warning("DBに該当月がないため新規登録になります: " + ", ".join(new_periods))

st.markdown(
    f"損益科目（コード4〜9）は月別発生額を決算月（{closing_month}月）から累積し、"
    "貸借科目（コード1〜3）は各月の残高として登録します。借方・貸方は損益科目の"
    "当月発生額から作成します。前年同月値は既存DB行から引き継ぎ、新規科目は0です。"
)
confirmed = st.checkbox(
    "関与先・対象年月・差額を確認しました。選択月の既存データを置き換えます。"
)

if st.button("選択した月を再登録", type="primary", disabled=not confirmed):
    completed = []
    errors = []
    progress = st.progress(0.0)

    for index, key in enumerate(selected_keys, start=1):
        _, month_frame = snapshots[key]
        year, month = divmod(key, 100)
        period_label = f"R{year - 2018}年{month}月" if year >= 2019 else month_label(key)
        wire = month_frame.rename(
            columns={
                "account_code": "勘定科目コード",
                "account_name": "勘定科目名",
                "previous_month": "前月末",
                "debit": "借方",
                "credit": "貸方",
                "current_month": "当月",
                "previous_year": "前年同月",
                "diff": "差額",
                "ratio": "対比",
            }
        )
        # The API removes byte-identical imports across periods; distinguish each monthly payload.
        wire["再登録対象年月"] = period_label
        payload = wire.to_csv(index=False).encode("utf-8-sig")

        try:
            response = requests.post(
                f"{API_BASE}/imports/trial-balance",
                data={"client_name": client_name, "period_label": period_label},
                files={"file": (uploaded.name, payload, "text/csv")},
                timeout=120,
            )
            response.raise_for_status()
            completed.append(month_label(key))
        except requests.RequestException as exc:
            errors.append(f"{month_label(key)}: {exc}")
        progress.progress(index / len(selected_keys))

    if completed:
        st.success("再登録しました: " + "、".join(completed))
    if errors:
        st.error("登録できなかった月: " + " / ".join(errors))
