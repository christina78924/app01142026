import streamlit as st
import pandas as pd
import numpy as np
import io

# =========================
# 全域設定
# =========================
st.set_page_config(page_title="IPQC CPK & Yield Generator", layout="wide")
st.title("📊 IPQC CPK & Yield 報表生成器")

CPK_LIMIT = 1.33

# =========================
# 站點順序
# =========================
TARGET_ORDER = [
    "MLA assy installation",
    "Mirror attachment",
    "Barrel attachment",
    "Condenser lens attach",
    "LED Module  attachment",
    "ILLU Module cover attachment",
    "Relay lens attachment",
    "LED FLEX GRAPHITE-1",
    "reflector attach",
    "singlet attach",
    "HWP Mylar attach",
    "PBS attachment",
    "Doublet attachment",
    "Top cover installation",
    "PANEL PRECISION AA（LAA）",
    "POST DAA INSPECTION",
    "PANEL FLEX ASSY",
    "LCOS GRAPHITE ATTACH",
    "DE OQC"
]

# =========================
# 工具函式
# =========================
def normalize_name(name):
    return (
        name.lower()
        .replace(" ", "")
        .replace("(", "")
        .replace(")", "")
        .replace("（", "")
        .replace("）", "")
        .replace("-", "")
        .replace("_", "")
    )

TARGET_MAP = {normalize_name(n): n for n in TARGET_ORDER}

def calculate_cpk(data, usl, lsl):
    data = pd.to_numeric(data, errors="coerce").dropna()
    if len(data) < 2:
        return np.nan
    mean = data.mean()
    std = data.std(ddof=1)
    if std == 0:
        return np.nan
    cpu = (usl - mean) / (3 * std) if not pd.isna(usl) else np.nan
    cpl = (mean - lsl) / (3 * std) if not pd.isna(lsl) else np.nan
    return min(cpu, cpl) if not pd.isna(cpu) and not pd.isna(cpl) else cpu if not pd.isna(cpu) else cpl

def excel_col_letter(idx):
    letters = ""
    while idx >= 0:
        idx, r = divmod(idx, 26)
        letters = chr(65 + r) + letters
        idx -= 1
    return letters

def highlight_low_cpk(val):
    try:
        if float(val) < CPK_LIMIT:
            return "background-color:#ffc7ce;color:#9c0006;font-weight:bold"
    except:
        pass
    return ""

# =========================
# Yield
# =========================
def process_yield(station, df):
    best_col, max_cnt = -1, 0
    for c in range(min(30, df.shape[1])):
        col = df.iloc[:, c].astype(str).str.upper()
        total = (col == "OK").sum() + (col == "NG").sum()
        if total > max_cnt:
            max_cnt, best_col = total, c
    if best_col == -1:
        return None
    col = df.iloc[:, best_col].astype(str).str.upper()
    ok = (col == "OK").sum()
    ng = (col == "NG").sum()
    return {
        "Station": station,
        "Total Qty": ok + ng,
        "OK Qty": ok,
        "NG Qty": ng,
        "Yield": ok / (ok + ng) if (ok + ng) else 0
    }

# =========================
# CPK
# =========================
def process_cpk(station, df):
    def find_row(keywords):
        for i in range(min(60, len(df))):
            row = " ".join(df.iloc[i].astype(str).str.lower())
            if any(k in row for k in keywords):
                return i
        return -1

    dim_row = find_row(["dim. no", "dim no"])
    usl_row = find_row(["usl"])
    lsl_row = find_row(["lsl"])
    if dim_row == -1:
        return []

    headers = df.iloc[dim_row].astype(str)
    ignore = {"date", "time", "no.", "remark", "judge", "note", "nan", ""}
    dim_cols = {i: h for i, h in enumerate(headers) if h.lower().strip() not in ignore and len(h.strip()) > 1}

    usls, lsls = {}, {}
    if usl_row != -1:
        for i, v in enumerate(df.iloc[usl_row]):
            try: usls[i] = float(v)
            except: pass
    if lsl_row != -1:
        for i, v in enumerate(df.iloc[lsl_row]):
            try: lsls[i] = float(v)
            except: pass

    start = max(dim_row, usl_row, lsl_row) + 1
    data = df.iloc[start:].copy()

    date_col = -1
    for c in range(min(15, data.shape[1])):
        if data.iloc[:, c].astype(str).str.contains(r"202\d-\d{2}-\d{2}").any():
            date_col = c
            break
    if date_col == -1:
        return []

    data["Date"] = data.iloc[:, date_col].astype(str).str.extract(r"(202\d-\d{2}-\d{2})")[0]
    data = data.dropna(subset=["Date"])

    results = []
    for date, g in data.groupby("Date"):
        for idx, dim in dim_cols.items():
            vals = g.iloc[:, idx]
            cpk = calculate_cpk(vals, usls.get(idx), lsls.get(idx))
            n = pd.to_numeric(vals, errors="coerce").dropna().size
            if n > 0:
                results.append({
                    "Station": station,
                    "Dim No": dim,
                    "Date": date,
                    "Sample Size": n,
                    "USL": usls.get(idx, ""),
                    "LSL": lsls.get(idx, ""),
                    "CPK": round(cpk, 3) if not pd.isna(cpk) else ""
                })
    return results

# =========================
# Streamlit 主流程
# =========================
uploaded = st.file_uploader("📂 上傳 Excel (.xlsx)", type=["xlsx"])

if uploaded:
    xls = pd.ExcelFile(uploaded)
    yield_list, cpk_list = [], []

    for sheet in xls.sheet_names:
        norm = normalize_name(sheet)
        station = next((v for k, v in TARGET_MAP.items() if k in norm), None)
        if not station:
            continue
        df = pd.read_excel(uploaded, sheet_name=sheet, header=None)
        y = process_yield(station, df)
        if y:
            yield_list.append(y)
        cpk_list.extend(process_cpk(station, df))

    df_yield = pd.DataFrame(yield_list)
    df_cpk = pd.DataFrame(cpk_list)

    if not df_yield.empty:
        df_yield["Station"] = pd.Categorical(df_yield["Station"], TARGET_ORDER, ordered=True)
        df_yield = df_yield.sort_values("Station")
        df_yield["Yield"] = (df_yield["Yield"] * 100).round(2).astype(str) + "%"

    if not df_cpk.empty:
        df_cpk["Station"] = pd.Categorical(df_cpk["Station"], TARGET_ORDER, ordered=True)
        df_cpk = df_cpk.sort_values(["Station", "Dim No", "Date"])

    tab1, tab2 = st.tabs(["Yield", "CPK"])

    with tab1:
        st.dataframe(df_yield, use_container_width=True)

    with tab2:
        st.write(df_cpk.style.applymap(highlight_low_cpk, subset=["CPK"]))

    # ===== Excel 匯出 =====
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_yield.to_excel(writer, sheet_name="Yield", index=False)
        df_cpk.to_excel(writer, sheet_name="CPK", index=False)

        ws = writer.sheets["CPK"]
        wb = writer.book
        cpk_col = excel_col_letter(df_cpk.columns.get_loc("CPK"))

        red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True})
        ws.conditional_format(
            f"{cpk_col}2:{cpk_col}{len(df_cpk)+1}",
            {"type": "cell", "criteria": "<", "value": CPK_LIMIT, "format": red}
        )

    output.seek(0)
    st.download_button(
        "📥 下載 Excel",
        output,
        "IPQC_Report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
