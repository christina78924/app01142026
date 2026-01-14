import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =========================
# Streamlit Page
# =========================
st.set_page_config(page_title="IPQC CPK & Yield", layout="wide")
st.title("📊 IPQC CPK & Yield 報表生成器")

# =========================
# Station Order
# =========================
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
    return str(name).lower().replace(" ","
