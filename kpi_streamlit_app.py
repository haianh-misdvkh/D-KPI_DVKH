#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
 DASHBOARD KPI & THƯỞNG CBNV - Khối Dịch vụ khách hàng (bản Streamlit)
============================================================================
Thay thế cho bản HTML tĩnh trước đây. App này TỰ ĐỘNG đọc lại dữ liệu mỗi
khi file "DATA_KPI.xlsx" (đặt cùng thư mục với file này) thay đổi - không
cần chạy script Python riêng để sinh lại báo cáo nữa.

CÁCH CHẠY
---------
    pip install -r requirements.txt      (chỉ cần 1 lần)
    streamlit run kpi_streamlit_app.py

CÁCH CẬP NHẬT DỮ LIỆU MỚI
--------------------------
    Chỉ cần THAY file "DATA_KPI.xlsx" bằng file mới (giữ nguyên tên), rồi
    bấm F5 / "🔄 Làm mới dữ liệu" trên app. Không cần chạy lại lệnh nào khác.

TÀI KHOẢN
---------
    Lưu trong "user_accounts.json" (tự tạo ở lần chạy đầu). Toàn bộ thay đổi
    Admin thực hiện (đổi mật khẩu, thêm/khoá/xoá user) được ghi thẳng vào
    file này trên máy chủ -> có hiệu lực ngay với TẤT CẢ mọi người, không
    cần export/import JSON thủ công như bản HTML cũ.
============================================================================
"""

import os
import re
import json
import math
import html as html_lib
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from pyxlsb import open_workbook as open_xlsb

# ============================================================================
# Cấu hình cố định
# ============================================================================
APP_DIR = Path(__file__).resolve().parent
DATA_FILE_CANDIDATES = ["DATA_KPI.xlsb", "DATA_KPI.xlsx"]
ACCOUNTS_FILE = APP_DIR / "user_accounts.json"

def find_data_file():
    for name in DATA_FILE_CANDIDATES:
        p = APP_DIR / name
        if p.exists():
            return p
    return APP_DIR / DATA_FILE_CANDIDATES[0]

DEFAULT_STAFF_PASSWORD = "123456"
DEFAULT_ADMIN_PASSWORD = "Admin@2026"

TEAM2_KEYWORDS = ["Hỗ trợ tổng đài Miền Bắc", "Hỗ trợ tổng đài Miền Nam"]
KH_COLUMNS_PROCESSED = ['Tên KH', 'Mã TC', 'POS_BOM', 'DPD_BOM', 'MAX_BUCKET', 'Tổng dư nợ', 'Dư ví', 'Tên nhân viên', 'Tổng thu', 'UNRESOLVED_MAX']
KH_COLUMNS_WALLET = ['Tên KH', 'Mã TC', 'POS_BOM', 'DPD_BOM', 'MAX_BUCKET', 'Tổng dư nợ', 'Dư ví', 'Tên nhân viên', 'Tổng thu']
KH_NUMERIC = {'POS_BOM', 'DPD_BOM', 'Tổng dư nợ', 'Dư ví', 'Tổng thu', 'UNRESOLVED_MAX'}
KH_DECIMALS = {'POS_BOM': 1, 'DPD_BOM': 0, 'Tổng dư nợ': 1, 'Dư ví': 1, 'Tổng thu': 1, 'UNRESOLVED_MAX': 0}
KH_PAGE_SIZE = 50

st.set_page_config(page_title="KPI & Thưởng CBNV - Khối DVKH", page_icon="🏦", layout="wide")

# ============================================================================
# CSS - giữ nguyên tông màu vàng đồng/kem của bản HTML gốc
# ============================================================================
st.markdown("""
<style>
  :root{
    --navy:#59451f; --navy-2:#776a51; --blue:#b98d34; --blue-dark:#8a5b00;
    --ink:#3a2f1c; --ink-2:#6b5d40; --muted:#8f8471; --line:#eee6d7;
    --good:#0ca30c; --good-bg:#e7f7e7; --warning:#c98500; --warning-bg:#fff6e0;
    --serious:#ec835a; --serious-bg:#fdeee7; --critical:#d03b3b; --critical-bg:#fbe9e9;
    font-variant-numeric: tabular-nums;
  }
  .block-container{padding-top:1rem; max-width:1280px;}
  .app-header{
    background: linear-gradient(110deg, var(--navy) 0%, var(--navy-2) 58%, #bda56d 100%);
    color:#fff; padding:16px 22px; border-radius:12px; margin-bottom:14px;
  }
  .app-header .brand{display:flex; align-items:center; gap:14px;}
  .app-header .brand-logo{width:40px;height:40px;border-radius:10px;background:rgba(255,255,255,.14);
    display:flex;align-items:center;justify-content:center;font-size:20px;flex:none;}
  .app-header h1{font-size:18px; margin:0; font-weight:700;}
  .app-header .sub{font-size:12px; color:#f2e8d3; margin-top:2px;}

  .avatar{width:38px;height:38px;border-radius:50%; background:var(--blue); color:#fff;
    display:flex; align-items:center; justify-content:center; font-weight:700; font-size:14px; flex:none;}
  .avatar.admin{background:var(--navy);}
  .name{font-weight:700; font-size:14.5px; color:var(--ink);}
  .role{font-size:11.5px; color:var(--ink-2);}
  .date-badge{display:inline-flex; align-items:center; gap:6px; background:#f4ead7; color:var(--blue-dark);
    padding:7px 12px; border-radius:20px; font-size:12px; font-weight:600; border:1px solid #e4d3a8; white-space:nowrap;}

  .status-pill{display:inline-flex; align-items:center; gap:5px; padding:4px 11px; border-radius:20px; font-size:12px; font-weight:700;}
  .status-good{background:var(--good-bg); color:var(--good);}
  .status-warning{background:var(--warning-bg); color:var(--warning);}
  .status-serious{background:var(--serious-bg); color:var(--serious);}
  .status-critical{background:var(--critical-bg); color:var(--critical);}
  .status-inactive{background:#f0f1f3; color:var(--muted);}

  .stat-grid{display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:12px; margin:10px 0 18px;}
  .stat-tile{border:1px solid var(--line); border-radius:10px; padding:14px 16px; background:#fffcf4;}
  .stat-tile .stat-label{font-size:11px; text-transform:uppercase; letter-spacing:.03em; color:var(--muted); font-weight:700; margin-bottom:6px;}
  .stat-tile .stat-value{font-size:21px; font-weight:800; color:var(--navy);}
  .stat-tile .stat-unit{font-size:11px; color:var(--muted); margin-top:2px;}

  .progress-track{width:100%; height:10px; border-radius:6px; background:#eceef1; overflow:hidden;}
  .progress-fill{height:100%; border-radius:6px; background:linear-gradient(90deg, var(--blue), var(--blue-dark));}

  .table-scroll{overflow-x:auto; border:1px solid var(--line); border-radius:10px; margin:8px 0 4px;}
  table{width:100%; border-collapse:collapse; font-size:13px;}
  thead th{background:#c6ad77; color:#43381f; text-align:left; padding:9px 10px; font-weight:700;
    border-bottom:1px solid var(--line); white-space:normal; word-break:break-word; line-height:1.28;
    text-transform:uppercase; letter-spacing:.02em; font-size:11.5px;}
  thead th.num{text-align:right;}
  tbody td{padding:9px 10px; border-bottom:1px solid #f2ece0; color:var(--ink); background:#fffdf7;}
  tbody tr:hover td{background:#faf4e8;}
  .num{text-align:right; font-variant-numeric:tabular-nums;}
  .total-row td{font-weight:800; background:#f4ead7 !important;}
  tfoot td{font-weight:800; background:#f4ead7; border-top:2px solid var(--blue);}
  .team-header-row td{background:var(--navy); color:#fff; font-weight:700; font-size:12.5px;}
  .kpi-detail-table td:first-child{white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:200px;}

  .note-box{background:#faf4e2; border-left:3px solid var(--blue); padding:10px 14px; border-radius:0 8px 8px 0;
    font-size:12px; color:#6b4f10; margin:8px 0 12px;}
  .warn-box{background:var(--warning-bg); border-left:3px solid var(--warning); padding:10px 14px; border-radius:0 8px 8px 0;
    font-size:12px; color:#7a5200; margin:8px 0 12px;}
  .badge-count{font-size:11.5px; background:#eef1f5; color:var(--ink-2); padding:3px 9px; border-radius:20px; font-weight:600; display:inline-block; margin-bottom:6px;}

  .role-badge{font-size:10.5px; padding:2px 8px; border-radius:20px; font-weight:700; text-transform:uppercase;}
  .role-admin{background:var(--navy); color:#fff;}
  .role-staff{background:#eef1f5; color:var(--ink-2);}
  .pwd-mono{font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f4f5f7; padding:2px 7px; border-radius:5px; font-size:12px;}
  .footer-note{text-align:center; padding:16px; font-size:11px; color:var(--muted);}

  .quick-nav{display:flex; gap:8px; flex-wrap:wrap; margin:2px 0 18px;}
  .quick-nav a{font-size:12.5px; color:var(--blue-dark); background:#f4ead7; padding:7px 14px; border-radius:20px;
    text-decoration:none; font-weight:700; border:1px solid #e4d3a8; cursor:pointer; white-space:nowrap;}
  .quick-nav a:hover{background:#eddfbf;}
  .sec-title{font-size:17px; font-weight:800; color:var(--navy); margin:6px 0 10px; display:flex; align-items:center; gap:8px; scroll-margin-top:64px;}

  /* ===== Responsive: tablet & mobile ===== */
  @media (max-width: 900px){
    .block-container{padding-left:0.7rem; padding-right:0.7rem;}
    .app-header{padding:12px 14px;}
    .app-header h1{font-size:15.5px; line-height:1.3;}
    .app-header .sub{font-size:11px;}
    .app-header .brand-logo{width:34px; height:34px; font-size:17px;}
    table{font-size:12px;}
    thead th{font-size:10.5px; padding:7px 8px; white-space:nowrap; word-break:normal;}
    tbody td, tfoot td{padding:7px 8px; white-space:nowrap;}
    .stat-tile{padding:11px 12px;}
    .stat-tile .stat-value{font-size:18px;}
    .quick-nav a{padding:6px 11px; font-size:12px;}
  }
  @media (max-width: 560px){
    .stat-grid{grid-template-columns:repeat(auto-fit,minmax(135px,1fr)); gap:8px;}
    .name{font-size:13.5px;}
    .role{font-size:10.5px;}
    .date-badge{font-size:11px; padding:6px 10px;}
    table{font-size:11px;}
    .avatar{width:32px; height:32px; font-size:12px;}
  }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# Hàm định dạng số (1 chữ số thập phân, kiểu Việt Nam: "." nghìn, "," thập phân)
# ============================================================================
def fmt_num(v, dec=1):
    if v is None or v == "":
        return "-"
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "-"
    if pd.isna(v):
        return "-"
    s = f"{v:,.{dec}f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return s

def fmt_pct(v, dec=1):
    if v is None or v == "":
        return "-"
    try:
        v = float(v) * 100
    except (TypeError, ValueError):
        return "-"
    if pd.isna(v):
        return "-"
    return fmt_num(v, dec) + "%"

def fmt_vnd(v):
    if v is None or v == "":
        return "-"
    formatted = fmt_num(v, 1)
    return formatted + " ₫" if formatted != "-" else "-"

def status_band(kpi):
    if kpi is None:
        return ("status-warning", "Chưa có dữ liệu")
    try:
        kpi = float(kpi)
    except (TypeError, ValueError):
        return ("status-warning", "Chưa có dữ liệu")
    if pd.isna(kpi):
        return ("status-warning", "Chưa có dữ liệu")
    if kpi >= 0.8: return ("status-good", "✅ Đạt tốt")
    if kpi >= 0.5: return ("status-warning", "⚠️ Khá")
    if kpi >= 0.3: return ("status-serious", "🔶 Cần cải thiện")
    return ("status-critical", "❌ Chưa đạt")

def initials(name):
    parts = (name or "?").strip().split(" ")
    return (parts[-1] or "?")[0].upper()

def esc(v):
    return html_lib.escape(str(v)) if v is not None else ""

# ============================================================================
# Đọc dữ liệu từ DATA_KPI.xlsx (giữ nguyên logic từ update_dashboard_v3.py)
# ============================================================================
def clean(s):
    if s is None:
        return None
    if isinstance(s, str):
        return re.sub(r'\s+', ' ', s).strip()
    return s

def num(v):
    if v is None or isinstance(v, str):
        return 0
    return v

def guess_chi_tieu_is_pct(ten, val):
    """Dùng khi không đọc được number_format (file .xlsb) - đoán % dựa vào tên chỉ tiêu / độ lớn giá trị."""
    name = (ten or '').lower()
    if 'thực thu' in name or 'thu ' in name or name.startswith('thu'):
        return False
    if val is None or val == "":
        return True
    try:
        v = float(val)
    except (TypeError, ValueError):
        return True
    return abs(v) <= 3

def read_xlsb_rows(file_path, sheet_name):
    """Đọc 1 sheet từ file .xlsb thành list các tuple giá trị, căn đúng theo số thứ tự dòng/cột
    (giống hệt cách openpyxl.iter_rows(values_only=True) trả về) để tái dùng logic parse phía dưới."""
    row_map = {}
    max_row = -1
    max_col = -1
    with open_xlsb(file_path) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            for row in sheet.rows():
                if not row:
                    continue
                r_idx = row[0].r
                d = {}
                for c in row:
                    d[c.c] = c.v
                    if c.c > max_col:
                        max_col = c.c
                row_map[r_idx] = d
                if r_idx > max_row:
                    max_row = r_idx
    rows = []
    for r in range(max_row + 1):
        d = row_map.get(r, {})
        rows.append(tuple(d.get(c) for c in range(max_col + 1)))
    return rows

def get_xlsb_sheet_names(file_path):
    with open_xlsb(file_path) as wb:
        return list(wb.sheets)

def parse_kpi_giao(rows, fmt_rows=None):
    total_rows_idx = [i for i, row in enumerate(rows) if row[2] is not None and str(row[2]).strip() == 'Total']

    employees = {}
    for n, idx in enumerate(total_rows_idx):
        row = rows[idx]
        user_id = row[1]
        if not user_id:
            continue
        emp = {
            "user_id": user_id, "stt": row[4], "ho_ten": clean(row[5]),
            "phong_bo_phan": clean(row[6]), "chuc_danh": clean(row[7]), "noi_lam_viec": row[8],
            "kpi_giao": {
                "du_no_giao": num(row[11]), "thuc_hien": num(row[14]), "kpi": num(row[16]),
                "kpi_tinh_thuong": num(row[17]), "muc_tien_thuong_co_so": num(row[18]),
                "incentive": num(row[21]), "thuong_chang": num(row[22]),
                "thuong_so_thu_xm_cu": num(row[23]), "thuong_thang": num(row[24]),
                "thuong_thi_dua_nvxs": num(row[25]), "tong_thuong": num(row[26]),
                "ghi_chu": row[27],
            },
            "chi_tieu_detail": []
        }
        end = total_rows_idx[n + 1] if n + 1 < len(total_rows_idx) else len(rows)
        for j in range(idx + 1, end):
            r = rows[j]
            if r[1] != user_id or r[9] is None:
                continue
            if fmt_rows is not None:
                chi_tieu_cell = fmt_rows[j][12] if len(fmt_rows[j]) > 12 else None
                chi_tieu_is_pct = bool(chi_tieu_cell is not None and '%' in (chi_tieu_cell.number_format or ''))
            else:
                chi_tieu_is_pct = guess_chi_tieu_is_pct(clean(r[9]), r[12])
            emp["chi_tieu_detail"].append({
                "ten": clean(r[9]), "trong_so": r[10], "du_no_giao": r[11], "chi_tieu": r[12],
                "chi_tieu_is_pct": chi_tieu_is_pct,
                "th_tinh_thuong": r[13], "thuc_hien": r[14], "pct_thuc_hien": r[15], "kpi": r[16],
            })
        employees[user_id] = emp
    return employees

def parse_overall(rows, employees):
    name_to_uid = {e['ho_ten']: uid for uid, e in employees.items()}

    report_date_raw = rows[0][1]
    m = re.search(r'(\d{2}/\d{2}/\d{4})', str(report_date_raw)) if report_date_raw else None
    report_date = m.group(1) if m else "N/A"

    header_idx1 = next((i for i, r in enumerate(rows) if r[1] == 'Row Labels'), None)
    if header_idx1 is not None:
        i = header_idx1 + 1
        while i < len(rows) and rows[i][1] and rows[i][1] != 'Grand Total':
            r = rows[i]
            uid = name_to_uid.get(r[1])
            if uid:
                employees[uid]["overall"] = {
                    "type": "manager", "team": None,
                    "du_no_giao": num(r[2]), "thuc_hien": num(r[3]), "kpi": num(r[4]), "tong_thuong": num(r[5]),
                }
            i += 1

    header_idx2 = None
    for i, r in enumerate(rows):
        if r[1] == 'Row Labels' and i != header_idx1:
            header_idx2 = i
            break
    if header_idx2 is not None:
        current_team = None
        i = header_idx2 + 1
        while i < len(rows) and rows[i][1] and rows[i][1] != 'Grand Total':
            r = rows[i]
            name = r[1]
            if r[2] is None:
                current_team = clean(name)
            else:
                uid = name_to_uid.get(name)
                if uid:
                    employees[uid]["overall"] = {
                        "type": "staff", "team": current_team,
                        "du_no_giao": num(r[2]), "thuc_hien": num(r[3]), "kpi": num(r[4]),
                        "kpi_tinh_thuong": num(r[5]), "incentive": num(r[6]), "chang": num(r[7]),
                        "thuong_thang": num(r[8]), "thuong_thi_dua_nvxs": num(r[9]),
                        "thuong_so_thu_xm_cu": num(r[10]), "tong_thuong": num(r[11]),
                    }
            i += 1
    return report_date

def to_id_str(v):
    if pd.isna(v):
        return None
    try:
        return str(int(v))
    except (ValueError, TypeError):
        return str(v)

def parse_customers(file_path, employees):
    filter_cols = ['Tên nhân viên', 'Tên nhân viên 2']
    all_cols = list(dict.fromkeys(filter_cols + KH_COLUMNS_PROCESSED + KH_COLUMNS_WALLET))
    engine = 'pyxlsb' if str(file_path).lower().endswith('.xlsb') else None
    df = pd.read_excel(file_path, sheet_name="Data", usecols=all_cols, engine=engine)

    df['POS_BOM'] = pd.to_numeric(df['POS_BOM'], errors='coerce').round(2)
    df['DPD_BOM'] = pd.to_numeric(df['DPD_BOM'], errors='coerce').round(0)
    df['Tổng dư nợ'] = pd.to_numeric(df['Tổng dư nợ'], errors='coerce').round(0)
    df['Dư ví'] = pd.to_numeric(df['Dư ví'], errors='coerce').round(0)
    df['Tổng thu'] = pd.to_numeric(df['Tổng thu'], errors='coerce').round(2)
    df['UNRESOLVED_MAX'] = pd.to_numeric(df['UNRESOLVED_MAX'], errors='coerce')
    df['Mã TC'] = df['Mã TC'].apply(to_id_str)
    df['MAX_BUCKET'] = df['MAX_BUCKET'].apply(lambda v: None if pd.isna(v) else str(v))

    def clean_row(r, cols):
        r = list(r)
        if 'UNRESOLVED_MAX' in cols:
            idx = cols.index('UNRESOLVED_MAX')
            if r[idx] is not None:
                r[idx] = int(r[idx])
        return r

    for uid, e in employees.items():
        team = e.get('phong_bo_phan', '') or ''
        use_col2 = any(k in team for k in TEAM2_KEYWORDS)
        col = 'Tên nhân viên 2' if use_col2 else 'Tên nhân viên'
        sub = df[df[col] == uid]

        processed = sub[(sub['UNRESOLVED_MAX'] == 1) & (sub['Tổng thu'].fillna(0) != 0)]
        wallet = sub[sub['Dư ví'].fillna(0) != 0]

        e["kh_processed"] = [clean_row(r, KH_COLUMNS_PROCESSED) for r in
                              processed[KH_COLUMNS_PROCESSED].replace({np.nan: None}).values.tolist()]
        e["kh_wallet"] = [clean_row(r, KH_COLUMNS_WALLET) for r in
                           wallet[KH_COLUMNS_WALLET].replace({np.nan: None}).values.tolist()]

@st.cache_data(show_spinner="Đang đọc dữ liệu từ file KPI...")
def load_all_data(file_path, _mtime):
    is_xlsb = str(file_path).lower().endswith('.xlsb')
    if is_xlsb:
        sheet_names_lower = [s.lower() for s in get_xlsb_sheet_names(file_path)]
        for required in ("KPI giao", "Overall", "Data"):
            if required.lower() not in sheet_names_lower:
                raise ValueError(f"File thiếu sheet bắt buộc: '{required}'")
        rows_kpi = read_xlsb_rows(file_path, "KPI giao")
        rows_overall = read_xlsb_rows(file_path, "Overall")
        employees = parse_kpi_giao(rows_kpi, fmt_rows=None)
        report_date = parse_overall(rows_overall, employees)
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for required in ("KPI giao", "Overall", "Data"):
            if required not in wb.sheetnames:
                raise ValueError(f"File thiếu sheet bắt buộc: '{required}'")
        ws1 = wb["KPI giao"]
        rows_kpi = list(ws1.iter_rows(min_row=1, max_row=ws1.max_row, values_only=True))
        fmt_rows = list(ws1.iter_rows(min_row=1, max_row=ws1.max_row))
        ws2 = wb["Overall"]
        rows_overall = list(ws2.iter_rows(min_row=1, max_row=ws2.max_row, values_only=True))
        employees = parse_kpi_giao(rows_kpi, fmt_rows=fmt_rows)
        report_date = parse_overall(rows_overall, employees)
    parse_customers(file_path, employees)
    return employees, report_date

def get_data():
    data_file = find_data_file()
    if not data_file.exists():
        return None, None
    mtime = data_file.stat().st_mtime
    return load_all_data(str(data_file), mtime)

# ============================================================================
# Tài khoản (lưu trực tiếp trên server -> có hiệu lực ngay, không cần export/import)
# ============================================================================
def load_accounts(employees):
    accounts = {}
    if ACCOUNTS_FILE.exists():
        try:
            accounts = json.loads(ACCOUNTS_FILE.read_text(encoding='utf-8'))
        except Exception:
            accounts = {}
    changed = False
    if "admin" not in accounts:
        accounts["admin"] = {"password": DEFAULT_ADMIN_PASSWORD, "role": "admin", "active": True,
                              "display_name": "Quản trị viên hệ thống"}
        changed = True
    for uid in employees.keys():
        if uid not in accounts:
            accounts[uid] = {"password": DEFAULT_STAFF_PASSWORD, "role": "staff", "active": True}
            changed = True
    if changed:
        save_accounts(accounts)
    return accounts

def save_accounts(accounts):
    ACCOUNTS_FILE.write_text(json.dumps(accounts, ensure_ascii=False, indent=2), encoding='utf-8')

# ============================================================================
# Session state
# ============================================================================
if 'auth' not in st.session_state:
    st.session_state.auth = False
    st.session_state.uid = None
    st.session_state.role = None

def do_logout():
    st.session_state.auth = False
    st.session_state.uid = None
    st.session_state.role = None

# ============================================================================
# Header (dùng chung mọi màn hình)
# ============================================================================
def render_header():
    st.markdown("""
    <div class="app-header">
      <div class="brand">
        <div class="brand-logo">🏦</div>
        <div>
          <h1>Khối Dịch vụ khách hàng - Tỷ lệ hoàn thành và thưởng CBNV</h1>
          <div class="sub">Báo cáo hiệu suất KPI &amp; thưởng cá nhân · Dữ liệu bảo mật theo tài khoản</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# Đăng nhập
# ============================================================================
def render_login(accounts):
    left, mid, right = st.columns([1, 1.3, 1])
    with mid:
        with st.container(border=True):
            st.markdown("### 🔐 Đăng nhập xem báo cáo")
            st.caption("Mỗi cán bộ chỉ xem được dữ liệu của chính mình")
            uid_in = st.text_input("Tên đăng nhập (user)", key="login_uid", placeholder="VD: cohd")
            st.caption("Là ký tự trong ngoặc đơn sau tên bạn, VD: *Hồ Đắc Cơ (cohd)* → user là **cohd**")
            pwd_in = st.text_input("Mật khẩu", type="password", key="login_pwd")

            if st.button("Đăng nhập", type="primary", use_container_width=True):
                uid = (uid_in or "").strip().lower()
                if not uid:
                    st.error("Vui lòng nhập tên đăng nhập.")
                elif uid not in accounts:
                    st.error("Không tìm thấy tài khoản này.")
                elif not accounts[uid].get('active', True):
                    st.error("Tài khoản đã bị khoá. Vui lòng liên hệ Admin.")
                elif pwd_in != accounts[uid]['password']:
                    st.error("Mật khẩu không đúng.")
                else:
                    st.session_state.auth = True
                    st.session_state.uid = uid
                    st.session_state.role = accounts[uid]['role']
                    st.rerun()

            with st.expander("🔑 Quên / đổi mật khẩu?"):
                fp_uid = st.text_input("Tên đăng nhập", key="fp_uid")
                fp_old = st.text_input("Mật khẩu hiện tại", type="password", key="fp_old")
                fp_new = st.text_input("Mật khẩu mới", type="password", key="fp_new", help="Tối thiểu 4 ký tự")
                fp_new2 = st.text_input("Nhập lại mật khẩu mới", type="password", key="fp_new2")
                if st.button("Xác nhận đổi mật khẩu", key="fp_submit"):
                    u = (fp_uid or "").strip().lower()
                    if u not in accounts:
                        st.error("Không tìm thấy tài khoản này.")
                    elif fp_old != accounts[u]['password']:
                        st.error("Mật khẩu hiện tại không đúng.")
                    elif len(fp_new) < 4:
                        st.error("Mật khẩu mới cần tối thiểu 4 ký tự.")
                    elif fp_new != fp_new2:
                        st.error("Xác nhận mật khẩu không khớp.")
                    else:
                        accounts[u]['password'] = fp_new
                        save_accounts(accounts)
                        st.success("Đổi mật khẩu thành công. Vui lòng đăng nhập lại với mật khẩu mới.")

            st.divider()
            st.caption("Mật khẩu mặc định lần đầu: **123456** — vui lòng đổi sau khi đăng nhập.")
            st.markdown('<div class="footer-note" style="padding:0;">made by Dungpv-QC_DVKH</div>', unsafe_allow_html=True)

# ============================================================================
# Toolbar sau đăng nhập
# ============================================================================
def render_toolbar(e, uid, role, report_date, accounts):
    c1, c2 = st.columns([2.2, 2])
    with c1:
        name = e['ho_ten'] if e else uid
        if role == 'admin':
            role_line = "Quản lý tài khoản người dùng"
        else:
            role_line = " · ".join([x for x in [e.get('chuc_danh'), e.get('phong_bo_phan'), e.get('noi_lam_viec')] if x]) if e else "Tài khoản chưa có dữ liệu KPI"
        st.markdown(f"""
          <div style="display:flex;align-items:center;gap:12px;">
            <div class="avatar {'admin' if role=='admin' else ''}">{esc(initials(name))}</div>
            <div><div class="name">{esc(name)} ({esc(uid)})</div><div class="role">{esc(role_line)}</div></div>
          </div>
        """, unsafe_allow_html=True)
    with c2:
        b1, b2, b3, b4 = st.columns([1.3, 1, 1, 1])
        b1.markdown(f'<div class="date-badge">📅 Dữ liệu ngày: {esc(report_date or "-")}</div>', unsafe_allow_html=True)
        with b2.popover("🔑 Đổi MK"):
            old_pwd = st.text_input("Mật khẩu hiện tại", type="password", key="tb_old")
            new_pwd = st.text_input("Mật khẩu mới", type="password", key="tb_new")
            new_pwd2 = st.text_input("Nhập lại mật khẩu mới", type="password", key="tb_new2")
            if st.button("Xác nhận", key="tb_submit"):
                if old_pwd != accounts[uid]['password']:
                    st.error("Mật khẩu hiện tại không đúng.")
                elif len(new_pwd) < 4:
                    st.error("Mật khẩu mới cần tối thiểu 4 ký tự.")
                elif new_pwd != new_pwd2:
                    st.error("Xác nhận mật khẩu không khớp.")
                else:
                    accounts[uid]['password'] = new_pwd
                    save_accounts(accounts)
                    st.success("Đã đổi mật khẩu.")
        if b3.button("🔄 Làm mới dữ liệu"):
            st.cache_data.clear()
            st.rerun()
        if b4.button("Đăng xuất", type="secondary"):
            do_logout()
            st.rerun()
    st.divider()

# ============================================================================
# Dashboard 1 - KPI
# ============================================================================
def render_kpi_section(e):
    k = e['kpi_giao']
    band_cls, band_label = status_band(k.get('kpi'))
    st.markdown(f'<span class="status-pill {band_cls}">{band_label}</span>', unsafe_allow_html=True)

    tiles = [
        ("Dư nợ giao", fmt_num(k.get('du_no_giao')), "triệu đồng"),
        ("$ Thực hiện", fmt_num(k.get('thuc_hien')), "theo chỉ tiêu giao"),
        ("KPI hoàn thành", fmt_pct(k.get('kpi')), "tổng hợp"),
        ("KPI tính thưởng", fmt_pct(k.get('kpi_tinh_thuong')), "sau quy đổi"),
    ]
    tiles_html = "".join(
        f'<div class="stat-tile"><div class="stat-label">{lbl}</div><div class="stat-value">{val}</div><div class="stat-unit">{unit}</div></div>'
        for lbl, val, unit in tiles)
    st.markdown(f'<div class="stat-grid">{tiles_html}</div>', unsafe_allow_html=True)

    pct = min(100, (k.get('kpi') or 0) * 100)
    st.markdown(f'''
      <div style="margin-bottom:10px;">
        <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--ink-2);margin-bottom:6px;">
          <span>Tiến độ hoàn thành KPI tổng</span><span>{fmt_pct(k.get("kpi"))}</span>
        </div>
        <div class="progress-track"><div class="progress-fill" style="width:{pct}%"></div></div>
      </div>''', unsafe_allow_html=True)

    rows_html = []
    for d in e['chi_tieu_detail']:
        chi_tieu_val = "-"
        if d.get('chi_tieu') is not None:
            chi_tieu_val = fmt_pct(d['chi_tieu']) if d.get('chi_tieu_is_pct') else fmt_num(d['chi_tieu'])
        rows_html.append(f'''<tr>
          <td title="{esc(d.get("ten") or "")}">{esc(d.get("ten") or "-")}</td>
          <td class="num">{fmt_pct(d["trong_so"]) if d.get("trong_so") is not None else "-"}</td>
          <td class="num">{fmt_num(d.get("du_no_giao"))}</td>
          <td class="num">{chi_tieu_val}</td>
          <td class="num">{fmt_num(d.get("th_tinh_thuong"))}</td>
          <td class="num">{fmt_num(d.get("thuc_hien"))}</td>
          <td class="num">{fmt_pct(d["pct_thuc_hien"]) if d.get("pct_thuc_hien") is not None else "-"}</td>
          <td class="num"><b>{fmt_pct(d["kpi"]) if d.get("kpi") is not None else "-"}</b></td>
          <td class="num">-</td><td class="num">-</td><td class="num">-</td>
        </tr>''')
    if not rows_html:
        rows_html = ['<tr><td colspan="11" style="text-align:center;color:var(--muted);padding:20px;">Không có dữ liệu chi tiết chỉ tiêu.</td></tr>']

    total_row = f'''<tr class="total-row">
      <td>Tổng cộng</td><td class="num">-</td>
      <td class="num">{fmt_num(k.get("du_no_giao"))}</td>
      <td class="num">-</td><td class="num">-</td>
      <td class="num">{fmt_num(k.get("thuc_hien"))}</td>
      <td class="num">-</td>
      <td class="num"><b>{fmt_pct(k.get("kpi"))}</b></td>
      <td class="num">{fmt_pct(k.get("kpi_tinh_thuong"))}</td>
      <td class="num">{fmt_num(k.get("muc_tien_thuong_co_so"))}</td>
      <td class="num">{fmt_num(k.get("tong_thuong"))}</td>
    </tr>'''

    table_html = f'''<div class="table-scroll"><table class="kpi-detail-table">
      <thead><tr>
        <th>Tên chỉ tiêu</th><th class="num">Trọng số</th><th class="num">Dư nợ giao (trđ)</th>
        <th class="num">Chỉ tiêu</th><th class="num">$ Thực hiện (tính thưởng)</th>
        <th class="num">$ Thực hiện</th><th class="num">% Thực hiện</th><th class="num">KPI</th>
        <th class="num">KPI tính thưởng</th><th class="num">Mức tiền thưởng cơ sở</th><th class="num">Tổng thưởng</th>
      </tr></thead>
      <tbody>{"".join(rows_html)}{total_row}</tbody>
    </table></div>'''
    st.markdown(table_html, unsafe_allow_html=True)

# ============================================================================
# Dashboard 2 - Thưởng
# ============================================================================
def render_thuong_section(e):
    o = e.get('overall')
    if not o:
        st.markdown('<div class="badge-count">Không có dữ liệu</div>', unsafe_allow_html=True)
        st.markdown('<div class="warn-box">Không tìm thấy dữ liệu thưởng tương ứng trong sheet Overall.</div>', unsafe_allow_html=True)
        return

    if o['type'] == 'manager':
        st.markdown('<div class="badge-count">Bảng cấp quản lý (Overall - bảng trên)</div>', unsafe_allow_html=True)
        st.markdown('<div class="note-box">Dữ liệu lấy từ bảng tổng hợp cấp quản lý (Row Labels) trong sheet Overall.</div>', unsafe_allow_html=True)
        html = f'''<div class="table-scroll"><table><thead><tr>
          <th>Row Labels</th><th class="num">Dư nợ giao (trđ)</th><th class="num">$ Thực hiện</th><th class="num">KPI</th><th class="num">Tổng thưởng</th>
        </tr></thead><tbody><tr class="total-row">
          <td>{esc(e["ho_ten"])}</td>
          <td class="num">{fmt_num(o.get("du_no_giao"))}</td>
          <td class="num">{fmt_num(o.get("thuc_hien"))}</td>
          <td class="num">{fmt_pct(o.get("kpi"))}</td>
          <td class="num">{fmt_vnd(o.get("tong_thuong"))}</td>
        </tr></tbody></table></div>'''
    else:
        st.markdown('<div class="badge-count">Bảng cán bộ (Overall - bảng dưới)</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="note-box">Dữ liệu lấy từ bảng chi tiết cán bộ trong sheet Overall, thuộc nhóm: {esc(o.get("team") or "-")}.</div>', unsafe_allow_html=True)
        html = f'''<div class="table-scroll"><table><thead><tr>
          <th>Row Labels</th><th class="num">Dư nợ giao (trđ)</th><th class="num">$ Thực hiện</th><th class="num">KPI</th>
          <th class="num">KPI tính thưởng</th><th class="num">Incentive</th><th class="num">Chặng</th>
          <th class="num">Thưởng Tháng</th><th class="num">Thưởng thi đua NVXS</th><th class="num">Thưởng Số thu XM cũ</th><th class="num">Tổng thưởng</th>
        </tr></thead><tbody>
        <tr class="team-header-row"><td colspan="11">{esc(o.get("team") or "Nhóm")}</td></tr>
        <tr class="total-row">
          <td>{esc(e["ho_ten"])}</td>
          <td class="num">{fmt_num(o.get("du_no_giao"))}</td>
          <td class="num">{fmt_num(o.get("thuc_hien"))}</td>
          <td class="num">{fmt_pct(o.get("kpi"))}</td>
          <td class="num">{fmt_pct(o.get("kpi_tinh_thuong"))}</td>
          <td class="num">{fmt_vnd(o.get("incentive"))}</td>
          <td class="num">{fmt_num(o.get("chang"))}</td>
          <td class="num">{fmt_vnd(o.get("thuong_thang"))}</td>
          <td class="num">{fmt_vnd(o.get("thuong_thi_dua_nvxs"))}</td>
          <td class="num">{fmt_vnd(o.get("thuong_so_thu_xm_cu"))}</td>
          <td class="num">{fmt_vnd(o.get("tong_thuong"))}</td>
        </tr>
        </tbody></table></div>'''
    st.markdown(html, unsafe_allow_html=True)

# ============================================================================
# Dashboard 3 - Danh sách khách hàng (2 bảng, chỉ xem)
# ============================================================================
def build_kh_table_html(cols, rows, total_row=None):
    thead = "".join(f'<th class="{"num" if c in KH_NUMERIC else ""}">{esc(c)}</th>' for c in cols)
    body_rows = []
    for r in rows:
        tds = []
        for i, c in enumerate(cols):
            v = r[i] if i < len(r) else None
            if c in KH_NUMERIC:
                tds.append(f'<td class="num">{fmt_num(v, KH_DECIMALS.get(c, 1))}</td>')
            else:
                tds.append(f'<td>{"-" if v in (None, "") else esc(v)}</td>')
        body_rows.append("<tr>" + "".join(tds) + "</tr>")
    if not body_rows:
        body_rows = [f'<tr><td colspan="{len(cols)}" style="text-align:center;color:var(--muted);padding:20px;">Không có dữ liệu.</td></tr>']
    tfoot = ""
    if total_row is not None:
        sum_thu, sum_vi = total_row
        tds = []
        for i, c in enumerate(cols):
            if i == 0:
                tds.append("<td>Tổng cộng</td>")
            elif c == 'Tổng thu':
                tds.append(f'<td class="num">{fmt_num(sum_thu, 1)}</td>')
            elif c == 'Dư ví':
                tds.append(f'<td class="num">{fmt_num(sum_vi, 1)}</td>')
            else:
                tds.append("<td></td>")
        tfoot = f"<tfoot><tr>{''.join(tds)}</tr></tfoot>"
    return f'<div class="table-scroll"><table><thead><tr>{thead}</tr></thead><tbody>{"".join(body_rows)}</tbody>{tfoot}</table></div>'

def render_kh_table(cols, rows, prefix):
    name_idx = cols.index('Tên KH')
    tong_thu_idx = cols.index('Tổng thu') if 'Tổng thu' in cols else -1
    du_vi_idx = cols.index('Dư ví') if 'Dư ví' in cols else -1

    sum_thu = sum((r[tong_thu_idx] or 0) for r in rows) if tong_thu_idx >= 0 else 0
    sum_vi = sum((r[du_vi_idx] or 0) for r in rows) if du_vi_idx >= 0 else 0

    query = st.text_input("🔍 Tìm theo tên khách hàng...", key=f"{prefix}_search", label_visibility="collapsed", placeholder="🔍 Tìm theo tên khách hàng...")
    prev_key = f"{prefix}_prev_query"
    page_key = f"{prefix}_page"
    if st.session_state.get(prev_key) != query:
        st.session_state[page_key] = 1
        st.session_state[prev_key] = query

    if query:
        q = query.strip().lower()
        filtered = [r for r in rows if q in str(r[name_idx] or "").lower()]
    else:
        filtered = rows

    total_pages = max(1, math.ceil(len(filtered) / KH_PAGE_SIZE))
    page = min(max(1, st.session_state.get(page_key, 1)), total_pages)
    st.session_state[page_key] = page

    start = (page - 1) * KH_PAGE_SIZE
    page_rows = filtered[start:start + KH_PAGE_SIZE]

    st.markdown(build_kh_table_html(cols, page_rows, total_row=(sum_thu, sum_vi)), unsafe_allow_html=True)

    pc1, pc2, pc3, pc4, pc5 = st.columns([3, 0.6, 0.7, 0.7, 0.6])
    pc1.caption(f"Trang {page}/{total_pages} · {fmt_num(len(filtered), 0)} bản ghi")
    if pc2.button("⏮", key=f"{prefix}_first", disabled=page <= 1, use_container_width=True):
        st.session_state[page_key] = 1; st.rerun()
    if pc3.button("‹ Trước", key=f"{prefix}_prev", disabled=page <= 1, use_container_width=True):
        st.session_state[page_key] = page - 1; st.rerun()
    if pc4.button("Sau ›", key=f"{prefix}_next", disabled=page >= total_pages, use_container_width=True):
        st.session_state[page_key] = page + 1; st.rerun()
    if pc5.button("⏭", key=f"{prefix}_last", disabled=page >= total_pages, use_container_width=True):
        st.session_state[page_key] = total_pages; st.rerun()

def render_kh_section(e):
    st.markdown(f'<div class="badge-count">{fmt_num(len(e["kh_processed"]), 0)} khách hàng</div>', unsafe_allow_html=True)
    st.markdown("#### 📗 Bảng 1 - Danh sách KH đã xử lý")
    st.markdown('<div class="note-box">Điều kiện: UNRESOLVED_MAX = 1 và Tổng thu khác 0. <i>Chỉ xem, không tải về.</i></div>', unsafe_allow_html=True)
    render_kh_table(KH_COLUMNS_PROCESSED, e['kh_processed'], "khp")

    st.divider()

    st.markdown(f'<div class="badge-count">{fmt_num(len(e["kh_wallet"]), 0)} khách hàng</div>', unsafe_allow_html=True)
    st.markdown("#### 📘 Bảng 2 - Danh sách KH có dư ví")
    st.markdown('<div class="note-box">Điều kiện: Dư ví khác 0. <i>Chỉ xem, không tải về.</i></div>', unsafe_allow_html=True)
    render_kh_table(KH_COLUMNS_WALLET, e['kh_wallet'], "khw")

# ============================================================================
# Trang Admin - quản lý tài khoản
# ============================================================================
def render_admin_panel(accounts, employees):
    st.markdown("### 👤 Quản lý tài khoản người dùng")

    col_a, col_b = st.columns([1, 1])
    with col_a:
        with st.popover("➕ Thêm user mới"):
            new_uid = st.text_input("Tên đăng nhập (user)", key="admin_add_uid", placeholder="VD: newuser01")
            new_name = st.text_input("Tên hiển thị (tuỳ chọn)", key="admin_add_name")
            new_pwd = st.text_input("Mật khẩu ban đầu", key="admin_add_pwd")
            if st.button("Lưu", key="admin_add_submit"):
                uid = (new_uid or "").strip().lower()
                if not uid:
                    st.error("Vui lòng nhập tên đăng nhập.")
                elif uid in accounts:
                    st.error("Tên đăng nhập đã tồn tại.")
                elif len(new_pwd) < 4:
                    st.error("Mật khẩu cần tối thiểu 4 ký tự.")
                else:
                    accounts[uid] = {"password": new_pwd, "role": "staff", "active": True}
                    if new_name.strip():
                        accounts[uid]["display_name"] = new_name.strip()
                    save_accounts(accounts)
                    st.success(f"Đã thêm tài khoản {uid}.")
                    st.rerun()
    with col_b:
        accounts_json = json.dumps(accounts, ensure_ascii=False, indent=2)
        st.download_button("📤 Tải file user_accounts.json (backup)", data=accounts_json,
                            file_name="user_accounts.json", mime="application/json", use_container_width=True)

    st.info("Đây là app chạy trên máy chủ (không còn là file HTML tĩnh) — mọi thay đổi ở đây được ghi thẳng vào "
            "file `user_accounts.json` trên server và có hiệu lực NGAY với TẤT CẢ cán bộ, không cần export/import thủ công.")

    query = st.text_input("🔍 Tìm theo user hoặc họ tên...", key="admin_search")
    q = (query or "").strip().lower()

    uids_sorted = sorted(accounts.keys(), key=lambda u: (u != 'admin', u))

    header_cols = st.columns([0.4, 1, 1.8, 0.9, 0.9, 1.2, 2.2])
    for c, label in zip(header_cols, ["#", "User", "Họ tên", "Vai trò", "Trạng thái", "Mật khẩu", "Hành động"]):
        c.markdown(f"**{label}**")

    i = 0
    for uid in uids_sorted:
        acc = accounts[uid]
        e = employees.get(uid)
        display_name = e['ho_ten'] if e else (acc.get('display_name') or "(Tài khoản thủ công — chưa có dữ liệu KPI)")
        if q and q not in uid.lower() and q not in display_name.lower():
            continue
        i += 1
        cols = st.columns([0.4, 1, 1.8, 0.9, 0.9, 1.2, 2.2])
        cols[0].write(i)
        cols[1].markdown(f"**{esc(uid)}**")
        cols[2].write(display_name)
        role_html = '<span class="role-badge role-admin">Admin</span>' if acc['role'] == 'admin' else '<span class="role-badge role-staff">Cán bộ</span>'
        cols[3].markdown(role_html, unsafe_allow_html=True)
        status_html = '<span class="status-pill status-good">Hoạt động</span>' if acc.get('active', True) else '<span class="status-pill status-inactive">Đã khoá</span>'
        cols[4].markdown(status_html, unsafe_allow_html=True)
        cols[5].markdown(f'<span class="pwd-mono">{esc(acc["password"])}</span>', unsafe_allow_html=True)

        with cols[6]:
            ac1, ac2, ac3 = st.columns(3)
            with ac1.popover("Đổi MK"):
                new_pwd2 = st.text_input("Mật khẩu mới", key=f"pwd_{uid}")
                if st.button("Xác nhận", key=f"pwdok_{uid}"):
                    if len(new_pwd2) < 4:
                        st.error("Tối thiểu 4 ký tự.")
                    else:
                        accounts[uid]['password'] = new_pwd2
                        save_accounts(accounts)
                        st.success("Đã đổi mật khẩu.")
                        st.rerun()
            if acc['role'] != 'admin':
                lbl = "Khoá" if acc.get('active', True) else "Mở khoá"
                if ac2.button(lbl, key=f"toggle_{uid}"):
                    accounts[uid]['active'] = not acc.get('active', True)
                    save_accounts(accounts)
                    st.rerun()
            can_delete = (not e) and acc['role'] != 'admin'
            if can_delete:
                del_confirm_key = f"del_confirm_{uid}"
                if not st.session_state.get(del_confirm_key):
                    if ac3.button("Xoá", key=f"del_{uid}"):
                        st.session_state[del_confirm_key] = True
                        st.rerun()
                else:
                    if ac3.button("⚠️ Xác nhận?", key=f"delok_{uid}"):
                        del accounts[uid]
                        save_accounts(accounts)
                        st.session_state[del_confirm_key] = False
                        st.rerun()

# ============================================================================
# MAIN
# ============================================================================
def main():
    render_header()

    data_file = find_data_file()
    if not data_file.exists():
        st.error(f"❌ Không tìm thấy file dữ liệu `DATA_KPI.xlsb` (hoặc `DATA_KPI.xlsx`). Vui lòng đặt file này cùng thư mục với app rồi bấm 'Rerun'.")
        st.stop()

    try:
        employees, report_date = get_data()
    except Exception as ex:
        st.error(f"❌ Lỗi khi đọc dữ liệu: {ex}")
        st.stop()

    accounts = load_accounts(employees)

    if not st.session_state.auth:
        render_login(accounts)
        st.markdown('<div class="footer-note">© 2026 Khối Dịch vụ khách hàng · made by Dungpv-QC_DVKH</div>', unsafe_allow_html=True)
        return

    uid = st.session_state.uid
    role = st.session_state.role
    if uid not in accounts:
        do_logout(); st.rerun()

    e = employees.get(uid)
    render_toolbar(e, uid, role, report_date, accounts)

    if role == 'admin':
        render_admin_panel(accounts, employees)
    else:
        if not e:
            st.markdown('<div class="warn-box">Tài khoản này chưa có dữ liệu KPI/Thưởng/Khách hàng tương ứng trong file dữ liệu hiện tại. Vui lòng liên hệ Admin nếu đây là tài khoản mới.</div>', unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="quick-nav">
              <a href="#sec-kpi">📈 Tỷ lệ hoàn thành KPI</a>
              <a href="#sec-thuong">💰 Thưởng</a>
              <a href="#sec-kh">👥 Danh sách khách hàng</a>
            </div>
            """, unsafe_allow_html=True)

            st.markdown('<div id="sec-kpi" class="sec-title">📈 Tỷ lệ hoàn thành KPI</div>', unsafe_allow_html=True)
            render_kpi_section(e)

            st.divider()
            st.markdown('<div id="sec-thuong" class="sec-title">💰 Thưởng</div>', unsafe_allow_html=True)
            render_thuong_section(e)

            st.divider()
            st.markdown('<div id="sec-kh" class="sec-title">👥 Danh sách khách hàng</div>', unsafe_allow_html=True)
            render_kh_section(e)

    st.markdown(f'<div class="footer-note">© 2026 Khối Dịch vụ khách hàng · Báo cáo tự động từ dữ liệu KPI ngày {esc(report_date or "-")}'
                f'<br>Dữ liệu bảo mật — chỉ hiển thị thông tin của tài khoản đã đăng nhập.<br>'
                f'<span style="opacity:.55;">made by Dungpv-QC_DVKH</span></div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
