# -*- coding: utf-8 -*-
"""
App tách đơn hàng Excel → kiểm tra → tạo SO → xuất Head/Line D365
Chạy: streamlit run app.py
"""

import importlib
from contextlib import contextmanager
from datetime import datetime

import pandas as pd
import streamlit as st

from config import OPTIONAL_FIELDS, REQUIRED_FIELDS
import utils.loader as _loader_mod
import utils.processor as _processor_mod
import utils.tax as _tax_mod
import utils.exporter as _exporter_mod

import utils.d365_merge as _d365_merge_mod

importlib.reload(_loader_mod)
importlib.reload(_tax_mod)
importlib.reload(_processor_mod)
importlib.reload(_exporter_mod)
importlib.reload(_d365_merge_mod)
from utils.loader import load_po_dataframe, load_po_template, load_vat_template
from utils.processor import build_vat_split_report, split_orders_by_po_and_item
from utils.exporter import (
    apply_sales_order_to_header,
    build_d365_export_filename,
    build_d365_files,
    fill_missing_site_warehouse,
    to_excel_bytes,
)
from utils.d365_merge import (
    build_mapping_preview,
    load_d365_header_file,
    merge_d365_header_with_lines,
    parse_d365_header_import,
)
from utils.validator import normalize_columns, validate_orders

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tách đơn D365",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* Font gốc 16px — toàn app */
    html, body, .stApp {
        font-size: 16px !important;
        font-family: "Segoe UI", Tahoma, sans-serif;
        line-height: 1.45;
    }
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li,
    label, .stCaption,
    button, input, textarea, select {
        font-size: 1rem !important;
    }
    /* Khong ep font len uploader */
    [data-testid="stFileUploader"] button,
    [data-testid="stFileUploader"] span,
    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] div {
        font-size: inherit;
    }
    [data-testid="stTabs"] button { font-size: 0.95rem !important; }
    [data-testid="stDataFrame"] div,
    [data-testid="stDataEditor"] div,
    [data-testid="stDataFrame"] td,
    [data-testid="stDataEditor"] td {
        font-size: 0.93rem !important;
    }
    /* Tiêu đề cột bảng — in đậm */
    [data-testid="stDataFrame"] [role="columnheader"],
    [data-testid="stDataEditor"] [role="columnheader"],
    [data-testid="stDataFrame"] thead th,
    [data-testid="stDataEditor"] thead th,
    [data-testid="stDataFrame"] th,
    [data-testid="stDataEditor"] th,
    [data-testid="stDataFrame"] .col-header,
    [data-testid="stDataFrame"] [class*="Header"],
    [data-testid="stDataFrame"] [class*="header-cell"],
    [data-testid="stDataEditor"] [class*="header-cell"] {
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    [data-testid="stDataFrame"] [role="columnheader"] *,
    [data-testid="stDataEditor"] [role="columnheader"] * {
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    /* Bảng HTML — tiêu đề cột đậm */
    .ql-df-wrap {
        overflow: auto;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        background: #fff;
        width: 100%;
    }
    table.ql-df {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.9rem;
        margin: 0;
    }
    table.ql-df thead th {
        font-weight: 700 !important;
        color: #0f172a !important;
        background: #f1f5f9 !important;
        border-bottom: 1px solid #cbd5e1;
        border-right: 1px solid #e2e8f0;
        padding: 0.45rem 0.55rem;
        text-align: left;
        white-space: nowrap;
        position: sticky;
        top: 0;
        z-index: 1;
    }
    table.ql-df tbody td {
        color: #334155;
        border-bottom: 1px solid #f1f5f9;
        border-right: 1px solid #f8fafc;
        padding: 0.35rem 0.55rem;
        white-space: nowrap;
    }
    table.ql-df tbody tr:hover td { background: #f8fafc; }
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] p {
        font-size: 0.95rem !important;
    }
    .stApp { background-color: #ffffff; }
    [data-testid="stAppViewContainer"] > .main {
        max-width: 100% !important;
        width: 100% !important;
    }
    [data-testid="stAppViewContainer"] .main .block-container,
    [data-testid="stMainBlockContainer"],
    .block-container {
        max-width: 100% !important;
        width: 100% !important;
        padding-top: 0.35rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 0.65rem !important;
        padding-right: 0.65rem !important;
    }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
        width: 100% !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.app-top-bar) {
        margin-top: 0.3rem;
        margin-bottom: 0.25rem;
        align-items: flex-end !important;
    }
    .app-top-bar {
        background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
        border-radius: 6px;
        padding: 0.4rem 0.85rem;
        color: #fff;
        display: flex;
        align-items: center;
        width: 100%;
        min-height: 2.6rem;
        box-sizing: border-box;
    }
    .app-top-bar .app-title {
        color: #fff !important;
        font-size: 1.05rem !important;
        font-weight: 600;
        margin: 0;
        line-height: 1.25;
    }
    .app-top-bar .app-flow {
        color: #dbeafe;
        margin: 0.08rem 0 0 0;
        font-size: 0.85rem !important;
        line-height: 1.25;
        opacity: 0.95;
    }
    .step-file-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.5rem;
        flex-wrap: wrap;
        margin: 0.1rem 0 0.25rem;
    }
    .step-file-row .file-tag { display: none !important; }
    .step-row {
        display: flex;
        align-items: center;
        gap: 0.15rem;
        flex-wrap: wrap;
        margin: 0;
    }
    .step-pill {
        background: #f1f5f9; border: 1px solid #e2e8f0;
        border-radius: 12px; padding: .2rem .55rem; font-size: .9rem !important; color: #475569;
    }
    .step-pill.active { background: #dbeafe; border-color: #93c5fd; color: #1d4ed8; font-weight: 600; }
    .step-pill.done   { background: #dcfce7; border-color: #86efac; color: #166534; }
    .step-arrow {
        color: #cbd5e1;
        font-size: 0.95rem !important;
        font-weight: 700;
        line-height: 1;
        padding: 0 0.1rem;
        user-select: none;
    }
    .step-arrow.done { color: #4ade80; }
    div[data-testid="stMetric"] {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 6px; padding: .3rem .5rem;
    }
    div[data-testid="stMetric"] label { font-size: .85rem !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { font-size: 1.15rem !important; }
    .section-title {
        font-size: 1rem !important; font-weight: 600; color: #1e293b;
        border-left: 3px solid #2563eb; padding-left: .4rem; margin: .4rem 0 .3rem;
    }
    .flow-box {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
        padding: .65rem .9rem; font-size: 1rem; color: #475569; margin-bottom: .6rem;
    }
    .map-ok {
        background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px;
        padding: .4rem .8rem; font-size: 1rem; color: #166534;
    }
    section[data-testid="stSidebar"] {
        background: #f8fafc;
    }
    section[data-testid="stSidebar"][aria-expanded="false"] {
        width: 0 !important;
        min-width: 0 !important;
    }
    section[data-testid="stSidebar"][aria-expanded="true"] {
        min-width: 380px !important;
        width: 380px !important;
    }
    section[data-testid="stSidebar"][aria-expanded="true"] > div:first-child {
        width: 380px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarContent"],
    section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        width: 100% !important;
        max-width: 100% !important;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 0.45rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        width: 100% !important;
    }
    section[data-testid="stSidebar"] h3 {
        font-size: 0.95rem !important;
        margin: 0.1rem 0 0.3rem 0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        font-size: 0.95rem !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button,
    section[data-testid="stSidebar"] div[data-testid="stDownloadButton"] > button {
        white-space: nowrap !important;
        font-size: 0.9rem !important;
        padding: 0.3rem 0.45rem !important;
        min-height: 2rem !important;
        line-height: 1.2 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
        padding: 0.25rem 0.5rem !important;
        min-height: 2.3rem !important;
        max-height: 2.5rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] svg {
        width: 1.1rem !important;
        height: 1.1rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] {
        font-size: 0.85rem !important;
        line-height: 1.15 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] small {
        display: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
        padding: 0.15rem 0.5rem !important;
        font-size: 0.85rem !important;
        min-height: 1.5rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stFileUploaderFileList"],
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] ul {
        min-height: 2.3rem !important;
        max-height: 2.5rem !important;
        padding: 0.25rem 0.5rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stCaption"] {
        font-size: 0.85rem !important;
    }
    section[data-testid="stSidebar"] hr {
        margin: 0.4rem 0 !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stExpander"] summary {
        padding-top: 0.25rem !important;
        padding-bottom: 0.25rem !important;
    }
    section[data-testid="stSidebar"] [data-testid="stDataFrame"],
    section[data-testid="stSidebar"] [data-testid="stDataEditor"] {
        width: 100% !important;
    }
    [data-testid="stFileUploader"] section {
        border-radius: 6px; border: 1.5px dashed #93c5fd !important; background: #eff6ff;
    }
    /* Khi đã có file: ẩn dropzone, hiện tên file trong cùng 1 hộp */
    [data-testid="stFileUploader"]:has([data-testid="stFileUploaderFile"]) section,
    [data-testid="stFileUploader"]:has(ul li) section {
        display: none !important;
    }
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFileList"],
    [data-testid="stFileUploader"] ul {
        margin: 0 !important;
        padding: 0.2rem 0.5rem !important;
        list-style: none !important;
        border: 1.5px dashed #93c5fd !important;
        border-radius: 6px !important;
        background: #eff6ff !important;
        min-height: 2.15rem !important;
        max-height: 2.15rem !important;
        display: flex !important;
        align-items: center !important;
        gap: 0.25rem !important;
        overflow: hidden !important;
    }
    [data-testid="stFileUploader"]:not(:has([data-testid="stFileUploaderFile"])):not(:has(ul li))
        [data-testid="stFileUploaderFileList"],
    [data-testid="stFileUploader"]:not(:has([data-testid="stFileUploaderFile"])):not(:has(ul li)) ul {
        display: none !important;
        border: none !important;
        padding: 0 !important;
        min-height: 0 !important;
        max-height: 0 !important;
    }
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
    [data-testid="stFileUploader"] li {
        margin: 0 !important;
        padding: 0 !important;
        min-height: 0 !important;
        width: 100% !important;
        display: flex !important;
        align-items: center !important;
        gap: 0.3rem !important;
        background: transparent !important;
        border: none !important;
    }
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"],
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] span,
    [data-testid="stFileUploader"] li span {
        font-size: 0.9rem !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        max-width: 100% !important;
    }
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFileData"] small,
    [data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] small,
    [data-testid="stFileUploader"] li small {
        display: none !important;
    }
    [data-testid="stFileUploader"] [data-testid="stFileUploaderDeleteBtn"] button,
    [data-testid="stFileUploader"] button[kind="icon"] {
        min-height: 1.2rem !important;
        height: 1.2rem !important;
        width: 1.2rem !important;
        padding: 0 !important;
    }
    .main [data-testid="stFileUploader"] section {
        padding: 0.2rem 0.5rem !important;
        min-height: 2.15rem !important;
        max-height: 2.15rem !important;
    }
    .main [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"],
    .main [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] span,
    .main [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] div {
        font-size: 0.9rem !important;
        line-height: 1.15 !important;
    }
    .main [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] small {
        display: none !important;
    }
    .main [data-testid="stFileUploader"] button {
        padding: 0.1rem 0.45rem !important;
        font-size: 0.85rem !important;
        min-height: 1.45rem !important;
        height: 1.45rem !important;
        line-height: 1 !important;
    }
    .main [data-testid="stFileUploader"] svg {
        width: 1rem !important;
        height: 1rem !important;
    }
    .step-file-row .file-tag { display: none !important; }
    /* Upload gọn trong các bước Xuất D365 (container có border) */
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stFileUploader"] section {
        padding: 0.2rem 0.45rem !important;
        min-height: 2.15rem !important;
        max-height: 2.15rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stFileUploader"] svg {
        width: 1rem !important;
        height: 1rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] {
        font-size: 0.9rem !important;
        line-height: 1.15 !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] span {
        font-size: 0.9rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stFileUploader"] button {
        padding: 0.1rem 0.45rem !important;
        font-size: 0.85rem !important;
        min-height: 1.45rem !important;
        height: 1.45rem !important;
        line-height: 1 !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 0.15rem;
        margin-bottom: 0.25rem;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        padding: 0.2rem 0.55rem !important;
        min-height: 0 !important;
        height: auto !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-panel"] {
        padding-top: 0.25rem !important;
    }
    div[data-testid="stExpander"] details { border-radius: 6px !important; }
    div[data-testid="stExpander"] summary {
        padding-top: 0.28rem !important;
        padding-bottom: 0.28rem !important;
        min-height: 0 !important;
        font-size: 0.95rem !important;
    }
    .workflow-step-title {
        font-size: 0.95rem !important;
        font-weight: 600;
        padding: 0.28rem 0.6rem;
        margin-top: 0.3rem;
        margin-bottom: 0;
        border-radius: 5px 5px 0 0;
        border: 1px solid #e2e8f0;
        border-bottom: none;
        background: #f1f5f9;
        color: #0f172a;
        letter-spacing: 0.01em;
    }
    .workflow-step-highlight {
        background: #eff6ff;
        border-color: #93c5fd;
        color: #1e40af;
    }
    .workflow-step-success {
        background: #ecfdf5;
        border-color: #86efac;
        color: #166534;
    }
    .accent-bar {
        height: 2px;
        width: 100%;
        margin: 0.2rem 0 0.25rem;
        border-radius: 999px;
        background: linear-gradient(90deg, #2563eb 0%, #60a5fa 45%, #bfdbfe 100%);
    }
    /* ── Tab Xuất D365: Summary / Actions / Detail ── */
    #export-tab-root { display: none; }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.export-section-mark) {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 8px !important;
        padding: 0.4rem 0.6rem 0.45rem !important;
        margin-bottom: 0.35rem !important;
    }
    .export-section-mark { display: none; }
    .export-section-label {
        font-size: 0.72rem !important;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        color: #475569;
        margin: 0 0 0.3rem 0;
        line-height: 1.2;
    }
    /* KPI dạng chip cùng 1 hàng: SO: 5 */
    .export-kpis {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: wrap !important;
        align-items: center;
        gap: 0.35rem;
        width: 100%;
        margin: 0 0 0.3rem;
        box-sizing: border-box;
    }
    .export-kpi {
        flex: 0 0 auto;
        min-width: 0;
        background: #f1f5f9 !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 999px;
        padding: 0.2rem 0.65rem;
        text-align: left;
        box-sizing: border-box;
        white-space: nowrap;
    }
    .export-kpi .kpi-label {
        display: inline !important;
        font-size: 0.8rem !important;
        color: #64748b !important;
        text-transform: none;
        letter-spacing: 0;
        margin: 0;
        font-weight: 500;
    }
    .export-kpi .kpi-value {
        display: inline !important;
        font-size: 0.85rem !important;
        font-weight: 700;
        color: #0f172a !important;
        line-height: 1.25;
        margin-left: 0.2rem;
    }
    .export-kpi.warn .kpi-value { color: #b45309 !important; }
    .export-kpi.ok .kpi-value { color: #15803d !important; }
    .export-alerts {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: wrap !important;
        align-items: center;
        gap: 0.35rem;
        width: 100%;
    }
    .export-status-ok {
        font-size: 0.8rem !important;
        color: #166534;
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 999px;
        padding: 0.18rem 0.55rem;
        margin: 0;
        line-height: 1.3;
        display: inline-block;
        flex: 0 0 auto;
        white-space: nowrap;
    }
    .export-vat-note {
        font-size: 0.8rem !important;
        color: #92400e;
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-radius: 999px;
        padding: 0.18rem 0.55rem;
        margin: 0;
        line-height: 1.3;
        display: inline-block;
        flex: 0 0 auto;
    }
    .export-status-wait {
        font-size: 0.82rem !important;
        color: #475569;
        background: #f8fafc;
        border: 1px dashed #cbd5e1;
        border-radius: 5px;
        padding: 0.25rem 0.5rem;
        margin: 0;
        line-height: 1.3;
    }
    .export-file-tag {
        font-size: 0.82rem !important;
        color: #065f46;
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-radius: 5px;
        padding: 0.28rem 0.55rem;
        margin: 0;
        line-height: 1.3;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .export-file-tag b { color: #047857; }
    /* Expander gọn cho nhóm Thao tác / Chi tiết */
    div[data-testid="stExpander"]:has(.export-fold-mark) details {
        border: 1px solid #e2e8f0 !important;
        border-radius: 8px !important;
        background: #fff !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark) summary {
        padding-top: 0.35rem !important;
        padding-bottom: 0.35rem !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
    }
    .export-fold-mark { display: none; }
    /* Nút action ngắn + upload gọn trong Thao tác */
    div[data-testid="column"]:has(.export-btn-anchor) div[data-testid="stDownloadButton"] {
        width: fit-content !important;
        max-width: 9rem !important;
        margin: 0 !important;
    }
    div[data-testid="column"]:has(.export-btn-anchor) div[data-testid="stDownloadButton"] > button {
        width: auto !important;
        min-width: 6.25rem !important;
        max-width: 9rem !important;
        min-height: 1.75rem !important;
        height: 1.75rem !important;
        padding: 0.12rem 0.55rem !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        white-space: nowrap !important;
        border-radius: 5px !important;
    }
    div[data-testid="stHorizontalBlock"]:has(.export-btn-anchor) {
        align-items: center !important;
        gap: 0.35rem !important;
        margin-bottom: 0.1rem !important;
    }
    /* Hàng nút Tải file — sát bảng, sát phải */
    .export-dl-bar { display: none; }
    div[data-testid="element-container"]:has(.export-dl-bar) {
        margin: 0 !important;
        padding: 0 !important;
        min-height: 0 !important;
    }
    div[data-testid="element-container"]:has(.export-dl-bar)
        + div[data-testid="element-container"] {
        margin-top: 0.15rem !important;
        margin-bottom: 0 !important;
        display: flex !important;
        justify-content: flex-end !important;
        width: 100% !important;
    }
    div[data-testid="element-container"]:has(.export-dl-bar)
        + div[data-testid="element-container"]
        div[data-testid="stDownloadButton"] {
        margin-left: auto !important;
        margin-right: 0 !important;
        width: fit-content !important;
    }
    /* Nút Tải file đã ghép — nền xanh (đè primary đỏ) */
    div[data-testid="element-container"]:has(.export-dl-bar)
        + div[data-testid="element-container"]
        div[data-testid="stDownloadButton"] > button,
    div[data-testid="column"]:has(.export-btn-dl) div[data-testid="stDownloadButton"] > button {
        background-color: #2563eb !important;
        border-color: #1d4ed8 !important;
        color: #ffffff !important;
        min-width: 8.5rem !important;
        max-width: 11rem !important;
        min-height: 1.85rem !important;
        height: 1.85rem !important;
        padding: 0.15rem 0.7rem !important;
        font-size: 0.82rem !important;
        font-weight: 600 !important;
    }
    div[data-testid="element-container"]:has(.export-dl-bar)
        + div[data-testid="element-container"]
        div[data-testid="stDownloadButton"] > button:hover,
    div[data-testid="column"]:has(.export-btn-dl) div[data-testid="stDownloadButton"] > button:hover {
        background-color: #1d4ed8 !important;
        border-color: #1e40af !important;
        color: #ffffff !important;
    }
    div[data-testid="element-container"]:has(.export-dl-bar)
        + div[data-testid="element-container"]
        div[data-testid="stDownloadButton"] > button:disabled,
    div[data-testid="column"]:has(.export-btn-dl) div[data-testid="stDownloadButton"] > button:disabled {
        background-color: #93c5fd !important;
        border-color: #93c5fd !important;
        color: #eff6ff !important;
        opacity: 0.9;
    }
    /* Siết khoảng trống dưới bảng trong expander chi tiết */
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stTabs"] [data-baseweb="tab-panel"] {
        padding-bottom: 0 !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark) .ql-df-wrap {
        margin-bottom: 0 !important;
    }
    /* Upload trong expander Thao tác — thấp, xám, ẩn mô tả dài */
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] section {
        padding: 0.1rem 0.4rem !important;
        min-height: 1.75rem !important;
        max-height: 1.75rem !important;
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 5px !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] small {
        display: none !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"],
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] [data-testid="stFileUploaderDropzoneInstructions"] span {
        font-size: 0.78rem !important;
        line-height: 1.1 !important;
        color: #64748b !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] button {
        padding: 0.08rem 0.4rem !important;
        font-size: 0.75rem !important;
        min-height: 1.35rem !important;
        height: 1.35rem !important;
        line-height: 1 !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="stFileUploader"] svg {
        width: 0.9rem !important;
        height: 0.9rem !important;
    }
    div[data-testid="stExpander"]:has(.export-fold-mark)
        [data-testid="element-container"] {
        margin-bottom: 0.05rem !important;
    }
    .workflow-inline-title {
        font-size: 0.84rem !important;
        font-weight: 650;
        color: #0f172a;
        margin: 0;
        line-height: 1.2;
        padding: 0;
        white-space: nowrap;
    }
    .workflow-inline-title.highlight { color: #1d4ed8; }
    .workflow-inline-title.success { color: #15803d; }
    .export-panel { margin-bottom: 0; }
    .export-action-row { margin: 0; padding: 0; }
    .export-section-divider { display: none; }
    .export-flow, .export-meta, .export-step-caption, .export-status-row { display: none; }
    .getting-started {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem 1.15rem;
        margin-top: 0.4rem;
        color: #475569;
        font-size: 1rem;
        line-height: 1.55;
    }
    .getting-started h3 {
        margin: 0 0 0.5rem 0;
        color: #1e3a5f;
        font-size: 1.1rem;
    }
    .getting-started ol {
        margin: 0.4rem 0 0 1.15rem;
        padding: 0;
    }
    .getting-started li { margin-bottom: 0.35rem; }
    .getting-started .gs-tip {
        margin-top: 0.75rem;
        padding: 0.55rem 0.7rem;
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 6px;
        color: #1e40af;
        font-size: 0.95rem;
    }
    .getting-started .gs-req {
        margin-top: 0.65rem;
        padding: 0.5rem 0.7rem;
        background: #fff;
        border: 1px dashed #cbd5e1;
        border-radius: 6px;
        font-size: 0.95rem;
        color: #334155;
    }
    .getting-started code {
        background: #e2e8f0;
        padding: 0.05rem 0.3rem;
        border-radius: 3px;
        font-size: 0.9rem;
    }
    div[data-testid="column"]:has(.app-top-bar) {
        display: flex;
        align-items: flex-end;
    }
    div[data-testid="column"]:has(.app-top-bar) > div,
    div[data-testid="column"]:has([data-testid="stFileUploader"]) > div {
        width: 100%;
    }
    div[data-testid="column"]:has([data-testid="stFileUploader"]) {
        display: flex;
        align-items: flex-end;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

MAP_LABELS = {
    "po_number": "PO# → Customer requisition",
    "customer": "Mã KH → Customer account",
    "item": "Mã hàng → Item number",
    "quantity": "Số lượng → Quantity",
    "product_name": "Tên hàng → Text",
    "store_name": "Tên CH → Delivery name",
    "delivery_date": "Ngày giao → Requested receipt date",
    "notes": "Ghi chú (Memo)",
    "uom": "Đơn vị → Unit",
    "vendor_code": "Mã NCC",
    "vendor_name": "Tên NCC",
    "order_date": "Ngày đặt",
    "unit_price": "Đơn giá → Unit price",
    "warehouse": "Kho → Warehouse",
    "site": "Site",
}

# ── Mẫu PO Import_Template ────────────────────────────────────────────────────
PO_TEMPLATE_XLSX = "PO Import_Template.xlsx"
_sample_df = load_po_template()
_sample_bytes = to_excel_bytes(("PO Import", _sample_df))

# ── Sample data: check thuế (Item → VAT) ─────────────────────────────────────
TAX_CHECK_FIELDS = {
    "item": ["Item number", "Item", "Mã hàng", "Ma hang", "SKU"],
    "product_name": ["Product name", "Tên hàng", "Ten hang", "Product Name"],
    "uom": ["Unit", "UOM", "ĐVT", "DVT"],
    "vat": ["Vat", "VAT", "Thuế suất", "Thue suat", "Tax"],
}

TAX_CHECK_SAMPLE_FILENAME = "VAT_Template.xlsx"
VAT_TEMPLATE_XLSX = "VAT_Template.xlsx"
TAX_CHECK_SHEET_NAME = "Tax"
TAX_COLUMNS = ["Item number", "Product name", "Unit", "Vat"]
TAX_SESSION_KEY = "tax_df"
TAX_FIELD_TO_COLUMN = {
    "item": "Item number",
    "product_name": "Product name",
    "uom": "Unit",
    "vat": "Vat",
}


def _default_tax_df() -> pd.DataFrame:
    return _normalize_tax_df(load_vat_template())


def _prepare_tax_df_for_editor(df: pd.DataFrame) -> pd.DataFrame:
    """Chuẩn hóa cột để hiển thị/sửa — giữ dòng trống khi user thêm mới."""
    work = df.copy()
    rename_map = {}
    cols_lower = {str(c).strip().lower(): c for c in work.columns}
    for std_col, aliases in TAX_CHECK_FIELDS.items():
        target = TAX_FIELD_TO_COLUMN[std_col]
        for alias in aliases:
            key = alias.strip().lower()
            if key in cols_lower:
                rename_map[cols_lower[key]] = target
                break

    work = work.rename(columns=rename_map)
    for col in TAX_COLUMNS:
        if col not in work.columns:
            work[col] = ""

    work = work[TAX_COLUMNS].fillna("")
    for col in TAX_COLUMNS:
        work[col] = work[col].astype(str)
    return work.reset_index(drop=True)


def _normalize_tax_df(df: pd.DataFrame) -> pd.DataFrame:
    """Chuẩn hóa file thuế khi export — bỏ dòng trống hoàn toàn."""
    work = _prepare_tax_df_for_editor(df)
    non_empty = work.apply(lambda row: any(str(v).strip() for v in row), axis=1)
    return work.loc[non_empty].reset_index(drop=True)


def _tax_df_to_bytes(df: pd.DataFrame) -> bytes:
    export = _normalize_tax_df(df) if not df.empty else pd.DataFrame(columns=TAX_COLUMNS)
    if "Item number" in export.columns:
        export["Item number"] = export["Item number"].astype(str).str.strip()
    return to_excel_bytes((TAX_CHECK_SHEET_NAME, export))


def _get_tax_df() -> pd.DataFrame:
    _init_tax_session()
    return _normalize_tax_df(st.session_state[TAX_SESSION_KEY])


def _vat_split_warnings(df_lines: pd.DataFrame) -> list[dict]:
    """Canh bao khi cung PO co nhieu thue VAT -> se tach SO."""
    warnings: list[dict] = []
    if "vat" not in df_lines.columns:
        return warnings

    grouped = (
        df_lines.groupby(["po_number", "customer"])["vat"]
        .nunique()
        .reset_index(name="vat_count")
    )
    for _, row in grouped[grouped["vat_count"] > 1].iterrows():
        vats = df_lines[
            (df_lines["po_number"] == row["po_number"]) & (df_lines["customer"] == row["customer"])
        ]["vat"].unique()
        warnings.append(
            {
                "row": "-",
                "field": "vat",
                "message": (
                    f"PO '{row['po_number']}' (KH {row['customer']}) có {len(vats)} thuế khác nhau "
                    f"({', '.join(str(v) for v in vats)}). Sẽ tách thành nhiều SO."
                ),
                "level": "warning",
            }
        )
    return warnings


def _init_tax_session() -> None:
    if TAX_SESSION_KEY not in st.session_state:
        st.session_state[TAX_SESSION_KEY] = _default_tax_df().copy()


def _load_tax_upload(uploaded_file) -> None:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        loaded_tax = pd.read_csv(uploaded_file, encoding="utf-8-sig", dtype=str)
    else:
        loaded_tax = pd.read_excel(uploaded_file, engine="openpyxl")
    if loaded_tax.empty:
        st.warning("File thuế trống — giữ nguyên dữ liệu hiện tại.")
        return

    st.session_state[TAX_SESSION_KEY] = _normalize_tax_df(loaded_tax)
    st.session_state.tax_loaded_sig = f"{uploaded_file.name}:{uploaded_file.size}"
    st.success(f"Đã nạp {len(st.session_state[TAX_SESSION_KEY])} dòng từ **{uploaded_file.name}**")


def _render_tax_sidebar() -> None:
    _init_tax_session()

    t1, t2, t3 = st.columns(3, gap="small")
    with t1:
        st.download_button(
            "Mẫu",
            data=to_excel_bytes(("VAT", load_vat_template())),
            file_name=VAT_TEMPLATE_XLSX,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="tax_template_xlsx_dl",
            help="Tải file mẫu VAT",
        )
    with t2:
        st.download_button(
            "Tải",
            data=_tax_df_to_bytes(st.session_state[TAX_SESSION_KEY]),
            file_name=TAX_CHECK_SAMPLE_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="tax_sample_dl",
            help="Tải bảng thuế hiện tại",
        )
    with t3:
        if st.button("Reset", use_container_width=True, key="tax_reset", help="Đặt lại bảng thuế mặc định"):
            st.session_state[TAX_SESSION_KEY] = _default_tax_df().copy()
            st.session_state.pop("tax_loaded_sig", None)
            st.rerun()

    tax_upload = st.file_uploader(
        "Upload thuế",
        type=["csv", "xlsx", "xls"],
        key="tax_file_upload",
        label_visibility="collapsed",
        help="Upload file thuế VAT (.csv / .xlsx)",
    )
    if tax_upload is not None:
        upload_sig = f"{tax_upload.name}:{tax_upload.size}"
        if st.session_state.get("tax_loaded_sig") != upload_sig:
            try:
                _load_tax_upload(tax_upload)
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file thuế: {e}")

    n_tax = len(_normalize_tax_df(st.session_state[TAX_SESSION_KEY]))
    st.caption(f"Bảng thuế · {n_tax} dòng")
    st.session_state[TAX_SESSION_KEY] = st.data_editor(
        _prepare_tax_df_for_editor(st.session_state[TAX_SESSION_KEY]),
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        height=200,
        column_config={
            "Item number": st.column_config.TextColumn("Item"),
            "Product name": st.column_config.TextColumn("Tên hàng"),
            "Unit": st.column_config.TextColumn("ĐVT"),
            "Vat": st.column_config.TextColumn("VAT"),
        },
    )


def _clear_po_session() -> None:
    for key in (
        "df_header",
        "df_lines",
        "d365_header_loaded_sig",
        "d365_so_mapping",
        "df_lines_merged",
        "d365_import_df",
        "d365_header_raw",
        "d365_merge_line_errors",
        "d365_merge_warnings",
    ):
        st.session_state.pop(key, None)


def _get_split_result(normalized_df: pd.DataFrame, so_prefix: str) -> tuple[pd.DataFrame | None, pd.DataFrame | None, list[dict]]:
    """Tach SO mot lan, tra ve header/lines hoac loi VAT."""
    tax_df = _get_tax_df()
    df_header, df_lines, vat_errors = split_orders_by_po_and_item(
        normalized_df, so_prefix=so_prefix, tax_df=tax_df
    )
    vat_blocking = [e for e in vat_errors if e.get("level") != "warning"]
    if vat_blocking:
        return None, None, vat_blocking
    return df_header, df_lines, []


@contextmanager
def _workflow_step(title: str, *, highlight: bool = False, success: bool = False):
    """Tieu de buoc quy trinh xuat D365."""
    classes = ["workflow-step-title"]
    if highlight:
        classes.append("workflow-step-highlight")
    if success:
        classes.append("workflow-step-success")
    st.markdown(f'<div class="{" ".join(classes)}">{title}</div>', unsafe_allow_html=True)
    with st.container(border=True):
        yield


def _vat_report_display(df_header: pd.DataFrame) -> pd.DataFrame:
    report = build_vat_split_report(df_header)
    if report.empty:
        return report
    return report.rename(columns={
        "po_number": "PO#",
        "customer": "Mã KH",
        "store_name": "Cửa hàng",
        "vat_count": "Số thuế",
        "vat": "VAT",
        "so_number": "SO tạm",
        "item_count": "Số MH",
        "line_count": "Số dòng",
        "total_qty": "Tổng SL",
        "notes": "Ghi chú cảnh báo",
    })


def _render_getting_started() -> None:
    st.markdown(
        """
        <div class="getting-started">
            <h3>Bắt đầu sử dụng</h3>
            <ol>
                <li><b>Chuẩn bị</b> — mở sidebar (biểu tượng &gt; góc trái) để tải <b>Mẫu PO</b> và kiểm tra <b>bảng thuế VAT</b>.</li>
                <li><b>Upload PO</b> — kéo file Excel/CSV vào hộp bên phải (hoặc Browse files).</li>
                <li><b>Kiểm tra</b> — xem tab Kiểm tra; sửa lỗi nếu có (thiếu cột, thiếu VAT…).</li>
                <li><b>Tách SO</b> — hệ thống tách theo PO + Khách hàng + VAT.</li>
                <li><b>Xuất D365</b> — Xuất File Header → import D365 lấy Sales order → upload lại Header → tải file đã ghép (Header + Line).</li>
            </ol>
            <div class="gs-req">
                <b>Cột bắt buộc trong file PO:</b>
                PO# · Mã KH (Store code) · Mã hàng (Item) · Số lượng (Qty)
            </div>
            <div class="gs-tip">
                Mẹo: nếu chưa có file PO, bấm <b>Mẫu PO</b> trong sidebar → điền dữ liệu → upload lại vào đây.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_export_tab(df_header: pd.DataFrame, df_lines: pd.DataFrame) -> None:
    d365_head, d365_line = build_d365_files(df_header, df_lines)
    vat_split_report = _vat_report_display(df_header)
    vat_split_count = len(vat_split_report)
    so_count = len(df_header)
    line_count = len(df_lines)

    header_xlsx = to_excel_bytes(("Header", d365_head))
    header_name = build_d365_export_filename("Header")

    mapped = bool(st.session_state.get("d365_so_mapping")) and not st.session_state.get(
        "d365_merge_line_errors"
    )
    map_count = len(st.session_state.get("d365_so_mapping") or {})
    status_label = f"Đã ghép {map_count}" if mapped else "Chưa import"
    status_cls = "ok" if mapped else ""

    # Chuẩn bị dữ liệu ghép (nếu đã map)
    header_out = d365_head
    d365_line_final = d365_line
    final_xlsx = None
    final_name = ""
    final_sheet_names = "Header + Line"
    map_preview = pd.DataFrame()

    if mapped:
        merged_lines = st.session_state.get("df_lines_merged")
        if merged_lines is not None:
            mapping = st.session_state["d365_so_mapping"]
            d365_head_built, d365_line_final = build_d365_files(df_header, merged_lines)
            d365_head_built = apply_sales_order_to_header(d365_head_built, mapping)

            raw_header = st.session_state.get("d365_header_raw")
            if isinstance(raw_header, pd.DataFrame) and not raw_header.empty:
                header_out = fill_missing_site_warehouse(
                    raw_header.copy(), d365_head_built, overwrite=True
                )
                header_out = fill_missing_site_warehouse(
                    header_out, df_header, overwrite=True
                )
            else:
                header_out = d365_head_built

            d365_line_final = fill_missing_site_warehouse(
                d365_line_final, d365_head_built, overwrite=True
            )
            d365_line_final = fill_missing_site_warehouse(
                d365_line_final, df_header, overwrite=True
            )

            final_sheets: list[tuple[str, pd.DataFrame]] = [
                ("Header", header_out),
                ("Line", d365_line_final),
            ]
            if vat_split_count:
                final_sheets.append(("Bao cao VAT", vat_split_report))
            final_name = build_d365_export_filename("Header_Line", "Final")
            final_xlsx = to_excel_bytes(*final_sheets)
            final_sheet_names = " + ".join(name for name, _ in final_sheets)
            map_preview = build_mapping_preview(mapping, df_header)

    st.markdown('<div id="export-tab-root"></div>', unsafe_allow_html=True)

    # ═══ 1. SUMMARY — 1 hàng dạng SO: 5 ════════════════════════════════════
    alerts: list[str] = []
    if mapped:
        alerts.append(
            f'<span class="export-status-ok">✓ Đã map {map_count}/{so_count} SO</span>'
        )
    if vat_split_count:
        alerts.append(
            f'<span class="export-vat-note">'
            f"! {vat_split_count} SO tách VAT – xem cột Customer reference."
            f"</span>"
        )
    st.markdown(
        f"""
        <div class="export-kpis">
          <div class="export-kpi">
            <span class="kpi-label">SO:</span><span class="kpi-value">{so_count}</span>
          </div>
          <div class="export-kpi">
            <span class="kpi-label">Line:</span><span class="kpi-value">{line_count}</span>
          </div>
          <div class="export-kpi {"warn" if vat_split_count else ""}">
            <span class="kpi-label">Tách VAT:</span><span class="kpi-value">{vat_split_count}</span>
          </div>
          <div class="export-kpi {status_cls}">
            <span class="kpi-label">Trạng thái:</span><span class="kpi-value">{status_label}</span>
          </div>
          {"".join(alerts)}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ═══ 2. ACTIONS — Xuất Header + Upload ═════════════════════════════════
    # Sau khi import Header D365 thành công → tự rút gọn nhóm Thao tác
    action_title = (
        "Thao tác — Xuất Header / Upload D365"
        + ("  ·  Đã import" if mapped else "  ·  Chưa import Header D365")
    )
    with st.expander(action_title, expanded=not mapped):
        st.markdown('<div class="export-fold-mark"></div>', unsafe_allow_html=True)

        # Cùng 1 hàng: Xuất Header | Upload
        btn1, up1 = st.columns([0.7, 3.3], gap="small")
        with btn1:
            st.markdown('<div class="export-btn-anchor"></div>', unsafe_allow_html=True)
            st.download_button(
                "Xuất Header",
                data=header_xlsx,
                file_name=header_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False,
                type="secondary",
                key="export_header_xlsx",
            )
        with up1:
            d365_header_upload = st.file_uploader(
                "Header D365",
                type=["csv", "xlsx", "xls"],
                key="d365_header_reimport",
                label_visibility="collapsed",
                help="Upload Header đã có cột Sales order từ D365",
            )

        if d365_header_upload is not None:
            upload_sig = f"{d365_header_upload.name}:{d365_header_upload.size}"
            if st.session_state.get("d365_header_loaded_sig") != upload_sig:
                try:
                    raw_d365_header = load_d365_header_file(d365_header_upload)
                    import_df, parse_errors = parse_d365_header_import(raw_d365_header)
                    if parse_errors:
                        _error_table(parse_errors, "Lỗi đọc Header D365", "error")
                    else:
                        merged_lines, mapping, line_errors, map_warnings = merge_d365_header_with_lines(
                            df_lines, import_df, df_header
                        )
                        st.session_state["d365_header_loaded_sig"] = upload_sig
                        st.session_state["d365_so_mapping"] = mapping
                        st.session_state["df_lines_merged"] = merged_lines
                        st.session_state["d365_import_df"] = import_df
                        st.session_state["d365_header_raw"] = raw_d365_header
                        st.session_state["d365_merge_line_errors"] = line_errors
                        st.session_state["d365_merge_warnings"] = map_warnings
                        if line_errors:
                            st.error(f"Lỗi ghép {len(line_errors)} dòng Line")
                        else:
                            st.rerun()
                except Exception as e:
                    st.error(f"Không đọc được file Header D365: {e}")

        if st.session_state.get("d365_so_mapping"):
            _error_table(st.session_state.get("d365_merge_warnings", []), "Cảnh báo map SO", "warning")
            merge_errors = st.session_state.get("d365_merge_line_errors", [])
            if merge_errors:
                _error_table(merge_errors, "Lỗi ghép Line", "error")

    # ═══ 3. DETAILED DATA — tabs + bảng, nút tải dưới cùng bên phải ═══════
    detail_title = (
        f"Dữ liệu chi tiết — Header ({len(header_out)}) · Line ({len(d365_line_final)})"
        + (f" · VAT ({vat_split_count})" if vat_split_count else "")
        + (f" · Map SO ({map_count})" if map_count else "")
    )
    with st.expander(detail_title, expanded=True):
        st.markdown('<div class="export-fold-mark"></div>', unsafe_allow_html=True)

        tab_labels = [
            f"Header ({len(header_out)})",
            f"Line ({len(d365_line_final)})",
        ]
        if vat_split_count:
            tab_labels.append("Báo cáo VAT")
        tab_labels.append(f"Map SO ({map_count})" if map_count else "Map SO")

        detail_tabs = st.tabs(tab_labels)
        with detail_tabs[0]:
            _show_df(header_out, height=360)
        with detail_tabs[1]:
            _show_df(d365_line_final, height=360)

        tab_idx = 2
        if vat_split_count:
            with detail_tabs[tab_idx]:
                _show_df(vat_split_report, height=360)
            tab_idx += 1

        with detail_tabs[tab_idx]:
            if map_count and not map_preview.empty:
                _show_df(map_preview, height=min(48 + map_count * 34, 360))
            else:
                st.caption("Upload Header D365 để xem bảng map SO.")

        # Nút Tải file — sát dưới bảng, sát phải, nền xanh
        st.markdown('<div class="export-dl-bar"></div>', unsafe_allow_html=True)
        if mapped and final_xlsx is not None:
            st.download_button(
                "Tải file đã ghép",
                data=final_xlsx,
                file_name=final_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False,
                type="primary",
                key="export_header_line_final_xlsx",
            )
        else:
            st.download_button(
                "Tải file đã ghép",
                data=b"",
                file_name="pending.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False,
                type="primary",
                key="export_header_line_final_disabled",
                disabled=True,
            )


def _render_steps(active: int, file_name: str | None = None) -> None:
    steps = ["① Upload", "② Kiểm tra", "③ Tách SO", "④ Xuất D365"]
    parts: list[str] = []
    for i, label in enumerate(steps):
        cls = "step-pill done" if i < active else ("step-pill active" if i == active else "step-pill")
        parts.append(f'<span class="{cls}">{label}</span>')
        if i < len(steps) - 1:
            arrow_cls = "step-arrow done" if i < active else "step-arrow"
            parts.append(f'<span class="{arrow_cls}" aria-hidden="true">→</span>')
    st.markdown(
        f'<div class="step-file-row"><div class="step-row">{"".join(parts)}</div></div>',
        unsafe_allow_html=True,
    )


def _show_map_table(col_map: dict, *, compact: bool = False) -> None:
    if not col_map:
        st.warning("Chưa nhận diện được cột nào. Kiểm tra tên cột trong file Excel.")
        return

    rows = [{"Trường hệ thống": MAP_LABELS.get(k, k), "Cột Excel": v} for k, v in col_map.items()]
    row_height = 35
    max_height = 220 if compact else 280
    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
        height=min(40 + len(rows) * row_height, max_height),
        column_config={
            "Trường hệ thống": st.column_config.TextColumn(width="medium"),
            "Cột Excel": st.column_config.TextColumn(width="medium"),
        },
    )
    if "po_number" in col_map and not compact:
        st.markdown(
            f'<div class="map-ok">✅ <b>PO#</b> đọc từ cột <b>「{col_map["po_number"]}」</b> &nbsp;·&nbsp; '
            f'vd: <code>PO0118070</code></div>',
            unsafe_allow_html=True,
        )


def _show_df(df: pd.DataFrame, *, height: int = 360) -> None:
    """Hiển thị bảng với tiêu đề cột in đậm (HTML)."""
    if df is None or df.empty:
        st.caption("Không có dữ liệu.")
        return
    html = df.to_html(index=False, escape=True, classes="ql-df", border=0)
    st.markdown(
        f'<div class="ql-df-wrap" style="max-height:{int(height)}px">{html}</div>',
        unsafe_allow_html=True,
    )


def _error_table(errors: list[dict], title: str, level: str) -> None:
    if not errors:
        return
    df_err = pd.DataFrame(errors)[["row", "field", "message"]].rename(
        columns={"row": "Dòng", "field": "Trường", "message": "Mô tả"}
    )
    if level == "error":
        st.error(f"**{title}** — {len(errors)} lỗi")
    else:
        st.warning(f"**{title}** — {len(errors)} cảnh báo")
    st.dataframe(df_err, use_container_width=True, hide_index=True, height=min(40 + len(errors) * 35, 280))


def _process_upload(uploaded, so_prefix: str) -> None:
    po_sig = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("po_loaded_sig") != po_sig:
        _clear_po_session()
        st.session_state["po_loaded_sig"] = po_sig

    try:
        raw_df = load_po_dataframe(uploaded)
    except Exception as e:
        st.error(f"Không đọc được file PO: {e}")
        return

    if raw_df.empty:
        st.warning("File PO trống — không có dữ liệu để xử lý.")
        return

    normalized_df, all_errors = validate_orders(raw_df)
    _, col_map = normalize_columns(raw_df)
    blocking_errors = [e for e in all_errors if e.get("level") != "warning"]
    warnings = [e for e in all_errors if e.get("level") == "warning"]
    step_active = 1 if blocking_errors else 3

    _render_steps(step_active, uploaded.name)

    tab1, tab2, tab3, tab4 = st.tabs(["Dữ liệu", "Kiểm tra", "Tách SO", "Xuất D365"])

    with tab1:
        with st.expander(f"Ánh xạ cột ({len(col_map)} trường)", expanded=False):
            _show_map_table(col_map, compact=True)
        st.dataframe(raw_df, use_container_width=True, height=520, hide_index=True)

    with tab2:
        if not blocking_errors and not warnings:
            st.success("Dữ liệu hợp lệ — chuyển tab **Tách SO** hoặc **Xuất D365**.")
        elif not blocking_errors:
            st.success("Không có lỗi nghiêm trọng — có thể tiếp tục tách SO.")

        _error_table(blocking_errors, "Lỗi cần sửa", "error")
        _error_table(warnings, "Cảnh báo", "warning")

        if not blocking_errors:
            _, preview_lines, vat_errors = split_orders_by_po_and_item(
                normalized_df, so_prefix=so_prefix, tax_df=_get_tax_df()
            )
            vat_blocking = [e for e in vat_errors if e.get("level") != "warning"]
            if vat_blocking:
                _error_table(vat_blocking, "Lỗi tra cứu thuế VAT", "error")
                st.info("Bổ sung mã hàng thiếu vào **Bảng thuế VAT** ở sidebar, rồi tải lại.")
            else:
                _error_table(_vat_split_warnings(preview_lines), "Cảnh báo tách SO theo thuế", "warning")

    with tab3:
        if blocking_errors:
            st.error("Còn lỗi dữ liệu — xem tab Kiểm tra.")
        else:
            df_header, df_lines, vat_blocking = _get_split_result(normalized_df, so_prefix)
            if df_header is None:
                _error_table(vat_blocking, "Lỗi tra cứu thuế VAT", "error")
                st.error("Thiếu VAT — cập nhật bảng thuế ở sidebar.")
            else:
                st.session_state["df_header"] = df_header
                st.session_state["df_lines"] = df_lines

                vat_split_report = build_vat_split_report(df_header)
                vat_split_po_count = (
                    vat_split_report["po_number"].nunique() if not vat_split_report.empty else 0
                )

                s1, s2, s3, s4 = st.columns(4)
                s1.metric("Số PO", df_header["po_number"].nunique())
                s2.metric("Số SO", len(df_header))
                s3.metric("Dòng Line", len(df_lines))
                s4.metric("PO tách VAT", vat_split_po_count)

                if vat_split_po_count:
                    st.warning(
                        f"Có **{vat_split_report['po_number'].nunique()}** PO bị tách thành nhiều SO "
                        f"do khác thuế VAT — xem cột **Customer reference** khi xuất Header Excel."
                    )

                hdr_cols = ["so_number", "po_number", "customer", "vat", "item_count", "line_count", "total_qty", "notes"]
                if "store_name" in df_header.columns:
                    hdr_cols.insert(4, "store_name")

                with st.expander("Header SO", expanded=True):
                    st.dataframe(
                        df_header[hdr_cols].rename(columns={
                            "so_number": "SO tạm", "po_number": "PO#", "customer": "Mã KH",
                            "vat": "VAT", "store_name": "Cửa hàng", "item_count": "Số MH",
                            "line_count": "Số dòng", "total_qty": "Tổng SL", "notes": "Ghi chú cảnh báo",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

                ex1, ex2, ex3 = st.columns(3)
                header_export = df_header[hdr_cols].rename(columns={
                    "so_number": "SO tạm", "po_number": "PO#", "customer": "Mã KH",
                    "vat": "VAT", "store_name": "Cửa hàng", "item_count": "Số MH",
                    "line_count": "Số dòng", "total_qty": "Tổng SL", "notes": "Ghi chú cảnh báo",
                })
                line_export_cols = ["so_number", "line_number", "po_number", "item", "quantity", "vat"]
                for opt in ("uom", "product_name"):
                    if opt in df_lines.columns:
                        line_export_cols.append(opt)
                line_export = df_lines[line_export_cols].rename(columns={
                    "so_number": "SO tạm", "line_number": "Dòng",
                    "po_number": "PO#", "item": "Mã hàng", "vat": "VAT",
                    "quantity": "SL", "uom": "ĐVT", "product_name": "Tên hàng",
                })
                with ex1:
                    st.download_button(
                        "Xuất Header SO (Excel)",
                        data=to_excel_bytes(("Header SO", header_export)),
                        file_name="Header_SO_Tam.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="split_header_xlsx",
                    )
                with ex2:
                    st.download_button(
                        "Xuất Line (Excel)",
                        data=to_excel_bytes(("Lines", line_export)),
                        file_name="Lines_SO_Tam.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="split_lines_xlsx",
                    )
                with ex3:
                    if vat_split_po_count:
                        report_display = vat_split_report.rename(columns={
                            "po_number": "PO#",
                            "customer": "Mã KH",
                            "store_name": "Cửa hàng",
                            "vat_count": "Số thuế",
                            "vat": "VAT",
                            "so_number": "SO tạm",
                            "item_count": "Số MH",
                            "line_count": "Số dòng",
                            "total_qty": "Tổng SL",
                            "notes": "Ghi chú cảnh báo",
                        })
                        st.download_button(
                            "Báo cáo tách VAT (Excel)",
                            data=to_excel_bytes(("Bao cao tach VAT", report_display)),
                            file_name="Bao_cao_tach_SO_theo_VAT.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary",
                            key="vat_split_report_xlsx",
                        )

                if vat_split_po_count:
                    with st.expander("Báo cáo tách SO theo VAT", expanded=True):
                        report_display = vat_split_report.rename(columns={
                            "po_number": "PO#",
                            "customer": "Mã KH",
                            "store_name": "Cửa hàng",
                            "vat_count": "Số thuế",
                            "vat": "VAT",
                            "so_number": "SO tạm",
                            "item_count": "Số MH",
                            "line_count": "Số dòng",
                            "total_qty": "Tổng SL",
                            "notes": "Ghi chú cảnh báo",
                        })
                        st.dataframe(report_display, use_container_width=True, hide_index=True)

                with st.expander("Chi tiết Line", expanded=False):
                    display_cols = ["so_number", "line_number", "po_number", "item", "quantity", "vat"]
                    for opt in ("uom", "product_name"):
                        if opt in df_lines.columns:
                            display_cols.append(opt)
                    st.dataframe(
                        df_lines[display_cols].rename(columns={
                            "so_number": "SO tạm", "line_number": "Dòng",
                            "po_number": "PO#", "item": "Mã hàng", "vat": "VAT",
                            "quantity": "SL", "uom": "ĐVT", "product_name": "Tên hàng",
                        }),
                        use_container_width=True,
                        hide_index=True,
                        height=240,
                    )

    with tab4:
        if blocking_errors:
            st.error("Chưa thể xuất — còn lỗi dữ liệu.")
        else:
            df_header = st.session_state.get("df_header")
            df_lines = st.session_state.get("df_lines")
            if df_header is None or df_lines is None:
                df_header, df_lines, vat_blocking = _get_split_result(normalized_df, so_prefix)
                if df_header is None:
                    _error_table(vat_blocking, "Lỗi tra cứu thuế VAT", "error")
                else:
                    st.session_state["df_header"] = df_header
                    st.session_state["df_lines"] = df_lines

            if df_header is not None and df_lines is not None:
                _render_export_tab(df_header, df_lines)


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Cấu hình & thuế")
    st.caption("Tiền tố SO · mẫu PO · bảng VAT dùng khi tách đơn")
    with st.expander("Cấu hình", expanded=True):
        cfg1, cfg2 = st.columns([1.45, 1], gap="small")
        with cfg1:
            so_prefix = st.text_input("Tiền tố SO", value="SO", help="Tiền tố số SO tạm khi tách đơn")
        with cfg2:
            st.markdown('<div style="height:1.55rem"></div>', unsafe_allow_html=True)
            st.download_button(
                "Mẫu PO",
                data=_sample_bytes,
                file_name=PO_TEMPLATE_XLSX,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="sample_dl",
                help="Tải file mẫu PO Import để điền và upload",
            )

        st.markdown("**Bảng thuế VAT**")
        st.caption("Map Item → VAT. Thiếu mã hàng sẽ chặn tách SO.")
        _render_tax_sidebar()

        with st.expander("Quy tắc tách SO", expanded=False):
            st.markdown(
                "- Cùng **PO + KH + VAT** → 1 SO\n"
                "- Khác VAT trong cùng PO → tách nhiều SO\n"
                "- Xuất Header → D365 cấp Sales order → ghép Line"
            )

# ── Header + Upload ───────────────────────────────────────────────────────────
_title_col, _upload_col = st.columns([2.75, 1], gap="small")
with _title_col:
    st.markdown(
        """
        <div class="app-top-bar">
            <div>
                <p class="app-title">📦 Tách đơn hàng Excel → D365</p>
                <p class="app-flow">Upload PO → Kiểm tra → Tách SO → Xuất Header → Import D365 → Ghép Line</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with _upload_col:
    uploaded = st.file_uploader(
        "Upload PO",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
        help="Kéo thả hoặc chọn file PO (.csv / .xlsx). Cần cột: PO#, Mã KH, Item, Qty.",
    )

if uploaded is None:
    _render_steps(0)
    _render_getting_started()
else:
    _process_upload(uploaded, so_prefix)
