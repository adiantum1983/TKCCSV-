import pandas as pd
import numpy as np

def compute_financial_metrics(filepath):
    # 1. 生データの読み込み
    if isinstance(filepath, str):
        is_csv = filepath.lower().endswith(".csv")
        file_obj = filepath
    else:
        is_csv = getattr(filepath, "name", "").lower().endswith(".csv")
        try:
            filepath.seek(0)
        except AttributeError:
            pass
        file_obj = filepath

    if is_csv:
        try:
            raw_df = pd.read_csv(file_obj, encoding="utf-8", header=None)
        except UnicodeError:
            if not isinstance(file_obj, str): file_obj.seek(0)
            raw_df = pd.read_csv(file_obj, encoding="shift_jis", header=None)
    else:
        raw_df = pd.read_excel(file_obj, header=None)
        
    # ヘッダー行を探す
    header_idx = -1
    for i, row in raw_df.iterrows():
        if any(isinstance(val, str) and ("科目" in val or "勘定" in val) for val in row.values):
            header_idx = int(str(i)) if str(i).isdigit() else i
            break
            
    if isinstance(header_idx, int) and header_idx > 0:
        if not isinstance(filepath, str): 
            try:
                filepath.seek(0)
            except AttributeError:
                pass
        
        if is_csv:
            try:
                raw_df = pd.read_csv(file_obj, encoding="utf-8", header=header_idx)
            except UnicodeError:
                if not isinstance(file_obj, str): file_obj.seek(0)
                raw_df = pd.read_csv(file_obj, encoding="shift_jis", header=header_idx)
        else:
            raw_df = pd.read_excel(file_obj, header=header_idx)
            
    # カラム抽出
    cols = [str(c).strip() for c in raw_df.columns]
    raw_df.columns = cols
    
    code_col = raw_df.columns[0]
    name_col = raw_df.columns[1]
    raw_df[code_col] = raw_df[code_col].astype(str).str.strip()
    
    # 「残高」列(当期累計) と 「前年同月」または「前期」列(前期累計)の特定
    curr_col = raw_df.columns[5] if len(raw_df.columns) > 5 else None
    for c in cols: 
        if "残高" in c: curr_col = c
        
    prev_yr_col = raw_df.columns[7] if len(raw_df.columns) > 7 else None
    for c in cols:
        if "前年同月" in c or "前期" in c: prev_yr_col = c
        
    # 値取得用ヘルパー
    def get_val(code, is_prev=False):
        col = prev_yr_col if is_prev else curr_col
        if not col or col not in raw_df.columns: return 0
        matched = raw_df[raw_df[code_col] == str(code)]
        if not matched.empty:
            return pd.to_numeric(matched.iloc[0][col], errors="coerce") or 0
        return 0
        
    def get_val_by_name(keywords, is_prev=False):
        col = prev_yr_col if is_prev else curr_col
        if not col or col not in raw_df.columns: return 0
        total = 0
        for idx, row in raw_df.iterrows():
            name = str(row[name_col])
            if any(kw in name for kw in keywords):
                total += pd.to_numeric(row[col], errors="coerce") or 0
        return total

    # 各指標の計算関数を作成する
    def calc_metrics_for_period(is_prev):
        sales = get_val("4000", is_prev)
        gp = get_val("5000", is_prev)
        if gp == 0: gp = sales - get_val("5200", is_prev)
        sga = get_val("6100", is_prev)
        op = get_val("6000", is_prev)
        if op == 0: op = gp - sga
        
        personnel = get_val_by_name(["役員報酬", "給与", "給料", "賞与引当金繰入", "法定福利費", "福利厚生費", "雑給"], is_prev)
        
        cash = get_val("1101", is_prev)
        ar = get_val("1122", is_prev)
        inv = get_val("1120", is_prev)
        
        total_assets = get_val("1000", is_prev)
        current_assets = get_val("1100", is_prev)
        fixed_assets = get_val("1200", is_prev)
        
        current_liabs = get_val("2100", is_prev)
        interest_bearing_debt = get_val("2113", is_prev) + get_val("2212", is_prev)
        equity = get_val("3000", is_prev)
        
        int_exp = get_val("7511", is_prev)
        net_income = get_val("9111", is_prev)
        cogs = get_val("5200", is_prev)
        
        return {
            "sales": sales, "gp": gp, "op": op, "sga": sga, "personnel": personnel,
            "cash": cash, "ar": ar, "inv": inv,
            "total_assets": total_assets, "current_assets": current_assets, "fixed_assets": fixed_assets,
            "current_liabs": current_liabs, "equity": equity, "ib_debt": interest_bearing_debt,
            "int_exp": int_exp, "net_income": net_income, "cogs": cogs
        }

    # 当期と前期の基礎数値をそれぞれ取得
    curr = calc_metrics_for_period(False)
    prev = calc_metrics_for_period(True) if prev_yr_col else None
    
    # Helper functions
    def safe_div(n, d): return n / d if d else 0
    def pct(n, d): return safe_div(n, d) * 100
    def f_pct(v): return f"{v:.1f}%" if v is not None else "-"
    def f_num(v): return f"{v:,.0f}" if v is not None else "-"
    def f_days(v): return f"{v:.1f}日" if v is not None else "-"
    def f_months(v): return f"{v:.1f}ヶ月" if v is not None else "-"
    def f_times(v): return f"{v:.2f}回" if v is not None else "-"
    def f_years(v): return f"{v:.1f}年" if v is not None else "-"
    
    # ------------------
    # 月次指標 (異常検知)
    # ------------------
    # 累計金額で処理するため、「月商」を簡便に「累計÷12（仮）」ではなく「累計÷経過月」で出したいが、
    # 経過月が不明なため「累計ベースの回転期間」に変更する。
    # ユーザー要望：「累計の金額で前期と比較」→ 全ての数値をYTDで出す
    
    def gen_monthly_row(name, detail, point, calc_curr, calc_prev=None, fmt=f_pct):
        val_curr = calc_curr(curr)
        val_prev = calc_prev(prev) if prev and calc_prev else None
        
        diff_str = "-"
        if val_curr is not None and val_prev is not None and val_prev != 0:
            if fmt == f_pct or fmt == f_times:
                # 率の差分（ポイント）
                diff_str = f"{(val_curr - val_prev):+.1f} pt"
            elif fmt == f_days or fmt == f_months or fmt == f_years:
                diff_str = f"{(val_curr - val_prev):+.1f}"
            else:
                # 金額ベースの場合は増減率%
                diff_str = f"{(val_curr - val_prev)/abs(val_prev)*100:+.1f}%"
        
        return {
            "指標(累計ベース)": name, 
            "内容/計算式": detail, 
            "当期(累計)": fmt(val_curr) if val_curr is not None else "-", 
            "前期(累計)": fmt(val_prev) if val_prev is not None else "-", 
            "差異(前年比)": diff_str,
            "見るポイント": point
        }

    monthly_rows = [
        {"指標(累計ベース)": "① 収益性", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        
        gen_monthly_row("売上高の成長", "売上高", "成長 or 減速", 
                        lambda c: c["sales"], lambda p: p["sales"], f_num),
                        
        gen_monthly_row("売上総利益率", "売上総利益 ÷ 売上高", "原価・値引き異常", 
                        lambda c: pct(c["gp"], c["sales"]), lambda p: pct(p["gp"], p["sales"])),
                        
        gen_monthly_row("営業利益", "売上総利益 − 販管費", "本業の利益", 
                        lambda c: c["op"], lambda p: p["op"], f_num),
                        
        gen_monthly_row("営業利益率", "営業利益 ÷ 売上高", "利益体質", 
                        lambda c: pct(c["op"], c["sales"]), lambda p: pct(p["op"], p["sales"])),
                        
        {"指標(累計ベース)": "② コスト", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        gen_monthly_row("人件費率", "人件費(累計) ÷ 売上高", "固定費の膨張", 
                        lambda c: pct(c["personnel"], c["sales"]), lambda p: pct(p["personnel"], p["sales"])),
        gen_monthly_row("販管費率", "販管費(累計) ÷ 売上高", "経費の増加", 
                        lambda c: pct(c["sga"], c["sales"]), lambda p: pct(p["sga"], p["sales"])),

        {"指標(累計ベース)": "③ 資金効率", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        # 回転「期間」は累計額を用いると日数が狂うため、「回転率（回/期間）」か「日次換算(x365)」とする
        gen_monthly_row("売掛金回転日数（年換算）", "売掛金 ÷ (売上高累計÷365)", "回収遅延", 
                        lambda c: safe_div(c["ar"], safe_div(c["sales"], 365)), lambda p: safe_div(p["ar"], safe_div(p["sales"], 365)), f_days),
        gen_monthly_row("在庫回転日数（年換算）", "在庫 ÷ (売上原価累計÷365)", "滞留在庫", 
                        lambda c: safe_div(c["inv"], safe_div(c["cogs"], 365)), lambda p: safe_div(p["inv"], safe_div(p["cogs"], 365)), f_days),

        {"指標(累計ベース)": "④ キャッシュ", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        gen_monthly_row("現預金残高", "現金＋預金", "即時支払能力", 
                        lambda c: c["cash"], lambda p: p["cash"], f_num),
        # 資金月数は累計ベースの「平均月商」(売上÷12)で算出するアプローチ（概算）
        gen_monthly_row("資金月数（ヶ月換算）", "現預金 ÷ (売上高累計÷12)", "何か月持つか", 
                        lambda c: safe_div(c["cash"], safe_div(c["sales"], 12)), lambda p: safe_div(p["cash"], safe_div(p["sales"], 12)), f_months),
    ]

    # ------------------
    # 四半期指標 (体質分析)
    # ------------------
    quarterly_rows = [
        {"指標(累計ベース)": "① 収益性", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        gen_monthly_row("ROA", "営業利益 ÷ 総資産", "資産効率", 
                        lambda c: pct(c["op"], c["total_assets"]), lambda p: pct(p["op"], p["total_assets"])),
        gen_monthly_row("ROE", "当期純利益 ÷ 自己資本", "株主利益", 
                        lambda c: pct(c["net_income"], c["equity"]), lambda p: pct(p["net_income"], p["equity"])),
        gen_monthly_row("営業利益率", "営業利益 ÷ 売上高", "収益力（安定評価）", 
                        lambda c: pct(c["op"], c["sales"]), lambda p: pct(p["op"], p["sales"])),

        {"指標(累計ベース)": "② 安全性", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        gen_monthly_row("自己資本比率", "自己資本 ÷ 総資産", "倒産耐性", 
                        lambda c: pct(c["equity"], c["total_assets"]), lambda p: pct(p["equity"], p["total_assets"])),
        gen_monthly_row("流動比率", "流動資産 ÷ 流動負債", "短期支払能力", 
                        lambda c: pct(c["current_assets"], c["current_liabs"]), lambda p: pct(p["current_assets"], p["current_liabs"])),
        gen_monthly_row("当座比率", "当座資産 ÷ 流動負債", "即時支払能力", 
                        lambda c: pct(c["cash"] + c["ar"], c["current_liabs"]), lambda p: pct(p["cash"] + p["ar"], p["current_liabs"])),

        {"指標(累計ベース)": "③ 効率性", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        gen_monthly_row("総資産回転率", "売上高累計 ÷ 総資産", "資産の稼働率", 
                        lambda c: safe_div(c["sales"], c["total_assets"]), lambda p: safe_div(p["sales"], p["total_assets"]), f_times),
        gen_monthly_row("固定資産回転率", "売上高累計 ÷ 固定資産", "設備効率", 
                        lambda c: safe_div(c["sales"], c["fixed_assets"]), lambda p: safe_div(p["sales"], p["fixed_assets"]), f_times),

        {"指標(累計ベース)": "④ 債務返済", "内容/計算式": "", "当期(累計)": "", "前期(累計)": "", "差異(前年比)": "", "見るポイント": ""},
        # 営業CFは簡易的に 営業利益+減価償却等 だが、ここでは純利益で代用するか、営業利益をベースにする。
        # 累計期間の厳密なCFが取れないため、「営業利益」で代用して計算。
        gen_monthly_row("債務償還簡易年数", "有利子負債 ÷ 営業利益(累計)", "返済年数", 
                        lambda c: safe_div(c["ib_debt"], c["op"]), lambda p: safe_div(p["ib_debt"], p["op"]), f_years) if curr["op"] > 0 else {"指標(累計ベース)": "債務償還簡易年数", "内容/計算式": "有利子負債 ÷ 営業利益(累計)", "当期(累計)": "赤字のため測定不能", "前期(累計)": "-", "差異(前年比)": "-", "見るポイント": "返済年数"},
        gen_monthly_row("インタレストカバレッジ", "営業利益(累計) ÷ 支払利息", "利息支払余力", 
                        lambda c: safe_div(c["op"], c["int_exp"]), lambda p: safe_div(p["op"], p["int_exp"]), lambda x: f"{x:.1f}倍")
    ]
    
    return pd.DataFrame(monthly_rows), pd.DataFrame(quarterly_rows)
