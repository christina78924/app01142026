import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="IPQC CPK & Yield", layout="wide")
st.title("📊 IPQC CPK & Yield 報表生成器")

TARGET_ORDER = [
    "MLA assy installation","Mirror attachment","Barrel attachment",
    "Condenser lens attach","LED Module  attachment",
    "ILLU Module cover attachment","Relay lens attachment",
    "LED FLEX GRAPHITE-1","reflector attach","singlet attach",
    "HWP Mylar attach","PBS attachment","Doublet attachment",
    "Top cover installation","PANEL PRECISION AA（LAA）",
    "POST DAA INSPECTION","PANEL FLEX ASSY",
    "LCOS GRAPHITE ATTACH","DE OQC"
]

def normalize_name(name):
    return str(name).lower().replace(" ","").replace("（","").replace("）","").replace("(","").replace(")","")

TARGET_MAP = {normalize_name(n): n for n in TARGET_ORDER}

def calculate_cpk(data, usl, lsl):
    data = pd.to_numeric(data, errors="coerce").dropna()
    if len(data) < 2:
        return np.nan
    std = data.std(ddof=1)
    if std == 0:
        return np.nan
    mean = data.mean()
    cpu = (usl - mean)/(3*std) if not pd.isna(usl) else np.nan
    cpl = (mean - lsl)/(3*std) if not pd.isna(lsl) else np.nan
    return min(cpu, cpl) if not pd.isna(cpu) and not pd.isna(cpl) else cpu if not pd.isna(cpu) else cpl

uploaded = st.file_uploader("📂 上傳 IPQC Excel", type=["xlsx"])

if uploaded:
    xls = pd.ExcelFile(uploaded)
    yield_rows, cpk_rows = [], []

    for sheet in xls.sheet_names:
        norm = normalize_name(sheet)
        station = next((v for k,v in TARGET_MAP.items() if k in norm), None)
        if not station:
            continue

        df = pd.read_excel(uploaded, sheet_name=sheet, header=None)

        def find_row(key):
            for i in range(min(80,len(df))):
                if key in " ".join(df.iloc[i].astype(str).str.lower()):
                    return i
            return -1

        dim_row = find_row("dim")
        usl_row = find_row("usl")
        lsl_row = find_row("lsl")

        # ---------- Yield ----------
        best_col, max_cnt = -1, 0
        for c in range(min(30, df.shape[1])):
            col = df.iloc[:,c].astype(str).str.upper()
            cnt = (col=="OK").sum() + (col=="NG").sum()
            if cnt > max_cnt:
                max_cnt, best_col = cnt, c

        if best_col != -1 and max_cnt > 0:
            col = df.iloc[:,best_col].astype(str).str.upper()
            ok = (col=="OK").sum()
            ng = (col=="NG").sum()
            yield_rows.append({
                "Station": station,
                "Total Qty": ok+ng,
                "OK Qty": ok,
                "NG Qty": ng,
                "Yield": ok/(ok+ng)
            })

        # ---------- CPK ----------
        if dim_row == -1:
            continue

        dim_headers = df.iloc[dim_row]
        dims = {
            i:str(v).strip() for i,v in enumerate(dim_headers)
            if str(v).strip() and str(v).lower() not in ["dim","dimension","nan"]
        }

        usls, lsls = {}, {}
        if usl_row!=-1:
            for i,v in enumerate(df.iloc[usl_row]):
                try: usls[i]=float(v)
                except: pass
        if lsl_row!=-1:
            for i,v in enumerate(df.iloc[lsl_row]):
                try: lsls[i]=float(v)
                except: pass

        start = max(dim_row, usl_row, lsl_row)+1
        data = df.iloc[start:].copy()

        date_col = -1
        for c in range(min(15, data.shape[1])):
            if data.iloc[:,c].astype(str).str.contains(r"202\d-\d{2}-\d{2}").any():
                date_col=c
                break
        if date_col==-1:
            continue

        data["Date"] = data.iloc[:,date_col].astype(str).str.extract(r"(202\d-\d{2}-\d{2})")[0]
        data = data.dropna(subset=["Date"])

        for date,g in data.groupby("Date"):
            for idx,dim_no in dims.items():
                vals = g.iloc[:,idx]
                n = pd.to_numeric(vals,errors="coerce").dropna().size
                if n>1:
                    cpk = calculate_cpk(vals, usls.get(idx), lsls.get(idx))
                    cpk_rows.append({
                        "Station": station,
                        "Dim No": dim_no,
                        "Date": date,
                        "Sample Size": n,
                        "USL": usls.get(idx,""),
                        "LSL": lsls.get(idx,""),
                        "CPK": cpk
                    })

    df_yield = pd.DataFrame(yield_rows)
    df_cpk = pd.DataFrame(cpk_rows)

    st.subheader("📈 Yield Summary")
    st.dataframe(df_yield.assign(Yield=lambda d: (d["Yield"]*100).round(2).astype(str)+"%"))

    st.subheader("📉 CPK Detail (CPK < 1.33 紅色)")
    st.dataframe(
        df_cpk.style.applymap(
            lambda v: "background-color:#ff9999" if isinstance(v,(int,float)) and v<1.33 else "",
            subset=["CPK"]
        )
    )

    # ---------- Excel Export ----------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_yield.to_excel(writer, sheet_name="Yield Summary", index=False)
        df_cpk.to_excel(writer, sheet_name="CPK Detail", index=False)

    wb = load_workbook(output)
    ws = wb["CPK Detail"]
    red = PatternFill("solid", fgColor="FF6666")
    cpk_col = [c.value for c in ws[1]].index("CPK")+1

    for r in range(2, ws.max_row+1):
        try:
            if float(ws.cell(r,cpk_col).value) < 1.33:
                ws.cell(r,cpk_col).fill = red
        except:
            pass

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)

    st.download_button(
        "📥 下載 IPQC CPK & Yield Excel",
        final,
        "IPQC_Final_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
