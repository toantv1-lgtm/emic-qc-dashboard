from datetime import date, datetime
import io
import os
import re
import sqlite3
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ================= 1. CẤU HÌNH DASHBOARD & CSS GIAO DIỆN CHUYÊN NGHIỆP =================
st.set_page_config(
    page_title="EMIC QC Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed", # Đóng Sidebar vì đã chuyển bộ lọc ra Main
)

st.markdown(
    """
    <style>
        header[data-testid="stHeader"] { display: none !important; }
        
        .main .block-container, div[data-testid="stAppViewBlockContainer"] {
            padding-top: 0.8rem !important;
            padding-bottom: 1rem !important;
            padding-left: 1.2rem !important;
            padding-right: 1.2rem !important;
            max-width: 100% !important;
        }

        .stApp {
            background-color: #F8FAFC;
            font-family: "Arial", sans-serif;
        }

        /* Khung Lọc Dữ Liệu Đỉnh Trang */
        .filter-card {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            padding: 15px 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            margin-bottom: 20px;
        }

        /* Tabs Navigation Căn Giữa */
        .stTabs [data-baseweb="tab-list"] {
            justify-content: center !important;
            gap: 10px !important;
            background-color: transparent !important;
            padding: 5px 0px 15px 0px !important;
            border-bottom: 2px solid #E2E8F0 !important;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #94A3B8 !important;
            color: #FFFFFF !important;
            border-radius: 6px !important;
            padding: 8px 24px !important;
            border: none !important;
            font-weight: bold !important;
            font-size: 14px !important;
        }
        .stTabs [aria-selected="true"] {
            background-color: #3B82F6 !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 6px rgba(59, 130, 246, 0.4) !important;
        }
        
        /* Card Khung Bao Biểu Đồ */
        .chart-card {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            padding: 10px 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.03);
            margin-bottom: 12px;
        }

        div[data-testid="stVerticalBlock"] > div { gap: 0.3rem !important; }
    </style>
""",
    unsafe_allow_html=True,
)

# Bảng Màu Chuẩn Matplotlib Gốc 100%
COLOR_SUCCESS = "#10B981"  # Xanh lá (Đạt / HT)
COLOR_PRIMARY = "#3B82F6"  # Xanh dương (Lệnh / FT)
COLOR_DANGER = "#EF4444"  # Đỏ (Lỗi / Lệnh Chưa Xong)
COLOR_WARNING = "#F59E0B"  # Cam (Đặc nhượng / SL Chưa Xong)
COLOR_PURPLE = "#A855F7"  # Tím (BY Sample)
COLOR_TEXT = "#0F172A"

DISTINCT_COLORS = [
    "#3B82F6", "#10B981", "#F59E0B", "#EF4444", "#A855F7",
    "#06B6D4", "#F43F5E", "#84CC16", "#F97316", "#14B8A6",
]

# Cấu hình Matplotlib ẩn (Chỉ dùng xuất Excel)
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["Arial", "sans-serif"]
plt.rcParams["font.size"] = 9
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


def log_formatter(x, pos):
  if x <= 0: return "0"
  if x >= 1e6: return f"{x*1e-6:.1f}M"
  if x >= 1e3: return f"{x*1e-3:.0f}K"
  return f"{int(x)}"


DB_NAME = "Report_Database.db"


@st.cache_data(ttl=30)
def load_data(tu_date, den_date):
  if not os.path.exists(DB_NAME):
    return pd.DataFrame(), pd.DataFrame()

  tu_iso = tu_date.strftime("%Y-%m-%d 00:00:00")
  den_iso = den_date.strftime("%Y-%m-%d 23:59:59")
  conn = sqlite3.connect(DB_NAME)
  df_qa32 = pd.read_sql_query(
      "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <= ?",
      conn, params=(tu_iso, den_iso),
  )
  df_coois = pd.read_sql_query(
      "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt <= ?",
      conn, params=(tu_iso, den_iso),
  )
  conn.close()
  return df_qa32, df_coois


# ================= 2. HÀM TẠO EXCEL ĐÍNH KÈM HÌNH ẢNH SẮC NÉT =================
def generate_print_ready_excel(
    phan_he_code, title_clean, tu_date, den_date,
    df_monthly, df_plan, df_family, df_year,
    fig1_mpl, fig2_mpl, fig3_mpl, fig4_mpl,
):
  output = io.BytesIO()
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  font_company = Font(name="Arial", size=10, bold=True, color="1F4E79")
  font_title = Font(name="Arial", size=14, bold=True, color="000000")
  font_subtitle = Font(name="Arial", size=9, italic=True, color="595959")
  font_section = Font(name="Arial", size=11, bold=True, color="1F4E79")
  font_header = Font(name="Arial", size=9, bold=True, color="FFFFFF")
  font_data = Font(name="Arial", size=9)
  fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
  fill_zebra = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
  thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
  header_border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"), top=Side(style="medium", color="1F4E79"), bottom=Side(style="medium", color="1F4E79"))
  align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

  sheets_data = [
      ("TienDo_Thang", "1. TIẾN ĐỘ SẢN XUẤT THEO THÁNG", df_monthly, fig1_mpl, "I7"),
      ("TongQuan_KeHoach", "2. TỔNG QUAN CHỈ TIÊU KẾ HOẠCH", df_plan, fig2_mpl, "E7"),
      ("Dong_SanPham", "3. CHI TIẾT THEO DÒNG SẢN PHẨM", df_family, fig3_mpl, "D7"),
      ("SanLuong_CaNam", "4. TỔNG SẢN LƯỢNG CẢ NĂM MÃ ĐẦU 5", df_year, fig4_mpl, "D7"),
  ]

  for sheet_name, section_title, df_table, fig_obj, img_pos in sheets_data:
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws["A1"] = "TỔNG CÔNG TY THIẾT BỊ ĐIỆN EMIC - PHÒNG QUẢN LÝ CHẤT LƯỢNG (QC)"
    ws["A1"].font = font_company
    ws["A2"] = f"BÁO CÁO TỔNG HỢP DỮ LIỆU - {title_clean.upper()}"
    ws["A2"].font = font_title
    ws["A3"] = f"Thời gian: {tu_date.strftime('%d/%m/%Y')} - {den_date.strftime('%d/%m/%Y')} | Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A3"].font = font_subtitle
    ws["A5"] = section_title
    ws["A5"].font = font_section

    start_row = 7
    for col_idx, col_name in enumerate(df_table.columns, start=1):
      cell = ws.cell(row=start_row, column=col_idx, value=col_name)
      cell.font = font_header
      cell.fill = fill_header
      cell.alignment = align_center
      cell.border = header_border

    for row_idx, row_data in enumerate(df_table.values, start=start_row + 1):
      row_fill = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
      for col_idx, val in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx)
        col_name_lower = str(df_table.columns[col_idx - 1]).lower()
        if isinstance(val, (int, float, np.number)):
          if "%" in col_name_lower or "tỷ lệ" in col_name_lower:
            cell.value = float(val) / 100.0 if val > 1.0 else float(val)
            cell.number_format = "0.0%"
          else:
            cell.value = float(val)
            cell.number_format = "#,##0" if float(val).is_integer() else "#,##0.00"
          cell.alignment = align_right
        else:
          cell.value = str(val)
          cell.alignment = align_center if col_idx == 1 or len(str(val)) < 10 else align_left
        cell.font = font_data
        if row_fill.fill_type: cell.fill = row_fill
        cell.border = thin_border

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        if cell.row >= start_row and cell.value is not None:
          max_len = max(max_len, len(str(cell.value)))
      ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    if fig_obj is not None:
      buf = io.BytesIO()
      fig_obj.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="#FFFFFF")
      buf.seek(0)
      img = OpenpyxlImage(buf)
      ws.add_image(img, img_pos)

  wb.save(output)
  return output.getvalue()


# ================= 3. BỘ LỌC NGÀY CHÍNH GIỮA MÀN HÌNH CỰC RÕ NÉT =================
st.markdown('<div class="filter-card">', unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #0F172A; margin-top: 0;'>📅 THIẾT LẬP DẢI THỜI GIAN BÁO CÁO</h4>", unsafe_allow_html=True)
col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 1])
with col_f1:
  tu_date = st.date_input("Từ ngày:", date(datetime.now().year, 1, 1))
with col_f2:
  den_date = st.date_input("Đến ngày:", date.today())
with col_f3:
  st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
  if st.button("🔄 CẬP NHẬT DỮ LIỆU", use_container_width=True, type="primary"):
    st.cache_data.clear()
    st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

df_qa32, df_coois = load_data(tu_date, den_date)

# ================= 4. NAVIGATION TABS =================
tab_vat_tu, tab_co_khi, tab_tuti, tab_cong_to = st.tabs([
    "📋 Báo Cáo Vật Tư", "⚙️ Báo Cáo Cơ Khí", "🔌 Báo Cáo TU/TI", "⚡ Báo Cáo Công Tơ",
])

# ================= 5. HÀM CHUNG BÁO CÁO COOIS (PLOTLY KHÔI PHỤC MATPLOTLIB) =================
def render_coois_tab_layout(phan_he_code, title_text):
  df_sub = df_coois[df_coois["phan_he"] == phan_he_code] if not df_coois.empty else pd.DataFrame()
  if df_sub.empty:
    st.info(f"💡 Chưa có dữ liệu sản xuất cho phân hệ {title_text}.")
    return

  months_labels = [f"T{i}" for i in range(1, 13)]
  m_comp_qty, m_uncomp_qty = [0.0]*12, [0.0]*12
  m_tot_orders, m_uncomp_orders = [0]*12, [0]*12
  tot_qty_all, deliv_qty_all = 0.0, 0.0

  for _, r in df_sub.iterrows():
    try: m_idx = datetime.strptime(str(r["ngay_lenh_dt"]).split()[0], "%Y-%m-%d").month - 1
    except: m_idx = 0
    if not (0 <= m_idx < 12): m_idx = 0
    sl_t, sl_h = float(r["sl_tong"]), float(r["sl_ht"])
    uncomp_q = max(0.0, sl_t - sl_h)
    tot_qty_all += sl_t
    deliv_qty_all += sl_h
    m_comp_qty[m_idx] += sl_h
    m_uncomp_qty[m_idx] += uncomp_q
    m_tot_orders[m_idx] += 1
    if sl_h < sl_t: m_uncomp_orders[m_idx] += 1

  title_clean = clean_emoji(title_text)

  # CẤU HÌNH CHIỀU CAO ĐỒNG BỘ CHO MỌI BIỂU ĐỒ LÀ 380px ĐỂ VUÔNG VẮN NHƯ ẢNH GỐC
  PLOT_HEIGHT = 380

  # HÀNG 1: BIỂU ĐỒ SẢN XUẤT (2.1) & DONUT (1.0)
  col1, col2 = st.columns([2.1, 1.0])

  with col1:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    m_comp_orders = [m_tot_orders[i] - m_uncomp_orders[i] for i in range(12)]

    # Kỹ thuật offsetgroup & base mô phỏng hoàn hảo Bar Chart của Matplotlib
    fig1.add_trace(go.Bar(x=months_labels, y=m_comp_orders, name="Lệnh Hoàn Thành", marker_color=COLOR_PRIMARY, offsetgroup=1), secondary_y=False)
    fig1.add_trace(go.Bar(x=months_labels, y=m_uncomp_orders, name="Lệnh Chưa Xong", marker_color=COLOR_DANGER, offsetgroup=1, base=m_comp_orders), secondary_y=False)
    fig1.add_trace(go.Bar(x=months_labels, y=m_comp_qty, name="SL Hoàn Thành", marker_color=COLOR_SUCCESS, offsetgroup=2), secondary_y=True)
    fig1.add_trace(go.Bar(x=months_labels, y=m_uncomp_qty, name="SL Chưa Xong", marker_color=COLOR_WARNING, offsetgroup=2, base=m_comp_qty), secondary_y=True)

    # Hiển thị % Phía trên Cột
    for i in range(12):
      t_qty = m_comp_qty[i] + m_uncomp_qty[i]
      pct = (m_comp_qty[i] / t_qty * 100) if t_qty > 0 else 0
      if t_qty > 0:
        fig1.add_annotation(
            x=months_labels[i], y=t_qty * 1.05, text=f"{pct:.0f}%",
            showarrow=False, font=dict(color=COLOR_SUCCESS, size=11, family="Arial", weight="bold"), yref="y2", xshift=14
        )

    fig1.update_layout(
        title=dict(text=f"TIẾN ĐỘ SẢN XUẤT - {title_clean}", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5),
        barmode="group", bargap=0.3,
        margin=dict(l=30, r=30, t=50, b=90), # b=90 để chứa Legend ngang bằng Donut
        height=PLOT_HEIGHT, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
        legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5, font=dict(size=11, family="Arial")),
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Arial")
    )
    fig1.update_xaxes(showgrid=False, tickfont=dict(size=11, family="Arial"))
    fig1.update_yaxes(title_text="← Tổng Lệnh", title_font=dict(size=12, color=COLOR_PRIMARY, weight="bold"), secondary_y=False, showgrid=True, gridcolor="#F1F5F9")
    fig1.update_yaxes(title_text="Số Lượng Giao [Log] →", title_font=dict(size=12, color=COLOR_SUCCESS, weight="bold"), type="log", secondary_y=True, showgrid=False)

    st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

  with col2:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    # Phục dựng màu sắc và format text chuẩn Matplotlib
    text_labels = [f"<b>Hoàn thành</b><br>{pct_deliv:.1f}%<br>({int(deliv_qty_all):,})", f"<b>Chưa xong</b><br>{pct_rem:.1f}%<br>({int(rem_qty_all):,})"]
    fig2 = go.Figure(data=[go.Pie(
        labels=["Hoàn thành", "Chưa xong"], values=[deliv_qty_all, rem_qty_all], hole=0.6,
        marker=dict(colors=[COLOR_SUCCESS, COLOR_WARNING], line=dict(color='#FFFFFF', width=3)),
        text=text_labels, textinfo="text", textposition="outside",
        textfont=dict(size=12, family="Arial", color=[COLOR_SUCCESS, COLOR_DANGER]), # Chữ màu khớp ảnh gốc
        direction="clockwise", sort=False
    )])
    fig2.update_layout(
        title=dict(text="TỶ LỆ HOÀN THÀNH TỔNG QUAN", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5),
        margin=dict(l=30, r=30, t=50, b=90), # Đồng bộ margin b=90 với Chart trái
        height=PLOT_HEIGHT, paper_bgcolor="#FFFFFF", showlegend=False,
        annotations=[dict(text=f"<b>TỔNG KẾ HOẠCH</b><br><span style='font-size:16px'>{int(tot_qty_all):,}</span>", x=0.5, y=0.5, font_size=12, font_family="Arial", showarrow=False)]
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

  # HÀNG 2: BỘ LỌC DÒNG SP
  sub_5 = df_sub[df_sub["ma_tp"].astype(str).str.split(".").str[0].str.lstrip("0").str.startswith("5")].copy() if not df_sub.empty else pd.DataFrame()
  raw_fams = [str(x).strip() for x in sub_5["mat_prefix"].unique() if pd.notna(x) and str(x).strip() and str(x).strip().lower() not in ["none", "nan"]] if not sub_5.empty else []
  available_fams = ["Tất cả dòng sản phẩm"] + sorted(list(set(raw_fams)))
  sel_fam = st.selectbox("🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}")

  col3, col4 = st.columns([1, 1])

  # --- DƯỚI TRÁI: BIỂU ĐỒ SẢN LƯỢNG DÒNG SP ---
  with col3:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    m3_qty = [0.0]*12
    if not sub_5.empty:
      sub_5_df = sub_5.copy()
      sub_5_df["month"] = pd.to_datetime(sub_5_df["ngay_lenh_dt"], errors="coerce").dt.month
      sub_filtered = sub_5_df[sub_5_df["mat_prefix"] == sel_fam] if sel_fam != "Tất cả dòng sản phẩm" else sub_5_df
      for _, r in sub_filtered.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12: m3_qty[int(m_val) - 1] += float(r["sl_ht"])

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(x=months_labels, y=m3_qty, name="SL Sản Xuất", marker_color=COLOR_PRIMARY, text=[f"{int(v):,}" if v > 0 else "" for v in m3_qty], textposition="outside", textfont=dict(color=COLOR_PRIMARY, size=11, family="Arial", weight="bold")))
    fig3.update_layout(
        title=dict(text=f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5),
        margin=dict(l=30, r=30, t=50, b=50), height=PLOT_HEIGHT - 40, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", showlegend=False, bargap=0.3
    )
    fig3.update_xaxes(showgrid=False, tickfont=dict(size=11, family="Arial"))
    fig3.update_yaxes(title_text="SL Hoàn Thành [Log]", title_font=dict(size=12, color=COLOR_PRIMARY, weight="bold"), type="log", showgrid=True, gridcolor="#F1F5F9")
    st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

  # --- DƯỚI PHẢI: TỔNG SẢN LƯỢNG MÃ ĐẦU 5 ---
  with col4:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    if not sub_5.empty:
      summary_fams = sub_5.groupby("mat_prefix")[["sl_ht"]].sum().reset_index()
      fams_x = [str(val) for val in summary_fams["mat_prefix"].tolist() if pd.notna(val) and str(val).strip() and str(val).strip().lower() not in ["none", "nan"]]
      if not fams_x: fams_x, deliv_fams = ["Trống"], [0.0]
      else:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x, deliv_fams = summary_fams["mat_prefix"].tolist(), summary_fams["sl_ht"].values
    else: fams_x, deliv_fams = ["Không có SP"], [0.0]

    bar_colors = [DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))]
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(x=fams_x, y=deliv_fams, marker_color=bar_colors, text=[f"{int(v):,}" if v > 0 else "" for v in deliv_fams], textposition="outside", textfont=dict(color=bar_colors, size=11, family="Arial", weight="bold")))
    fig4.update_layout(
        title=dict(text="TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5),
        margin=dict(l=30, r=30, t=50, b=50), height=PLOT_HEIGHT - 40, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", showlegend=False, bargap=0.3
    )
    fig4.update_xaxes(showgrid=False, tickfont=dict(size=11, family="Arial", weight="bold"))
    fig4.update_yaxes(title_text="Số Lượng SP [Log]", title_font=dict(size=12, color=COLOR_SUCCESS, weight="bold"), type="log", showgrid=True, gridcolor="#F1F5F9")
    st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

  # ================= XUẤT ẢNH ẨN MATPLOTLIB CHO FILE EXCEL =================
  fig1_mpl, ax_m1 = plt.subplots(figsize=(8, 3.2), dpi=200)
  ax_m1_twin = ax_m1.twinx()
  ax_m1.bar(np.arange(12) - 0.18, m_comp_orders, width=0.35, color=COLOR_PRIMARY, label="Lệnh HT")
  ax_m1.bar(np.arange(12) - 0.18, m_uncomp_orders, width=0.35, bottom=m_comp_orders, color=COLOR_DANGER, label="Lệnh Chưa Xong")
  ax_m1_twin.bar(np.arange(12) + 0.18, m_comp_qty, width=0.35, color=COLOR_SUCCESS, label="SL HT")
  ax_m1_twin.bar(np.arange(12) + 0.18, m_uncomp_qty, width=0.35, bottom=m_comp_qty, color=COLOR_WARNING, label="SL Chưa Xong")
  ax_m1.set_title(f"TIẾN ĐỘ SẢN XUẤT - {title_clean}", fontweight="bold")
  ax_m1.set_xticks(np.arange(12))
  ax_m1.set_xticklabels(months_labels)
  ax_m1_twin.set_yscale("symlog", linthresh=100)
  plt.close(fig1_mpl)

  fig2_mpl, ax_m2 = plt.subplots(figsize=(4, 3.2), dpi=200)
  ax_m2.set_aspect("equal")
  if tot_qty_all > 0:
    ax_m2.pie([deliv_qty_all, rem_qty_all], labels=["Hoàn thành", "Chưa xong"], colors=[COLOR_SUCCESS, COLOR_WARNING], autopct="%1.1f%%", startangle=140, wedgeprops=dict(width=0.35, edgecolor="white"))
  ax_m2.set_title("TỶ LỆ HOÀN THÀNH TỔNG QUAN", fontweight="bold")
  plt.close(fig2_mpl)

  fig3_mpl, ax_m3 = plt.subplots(figsize=(6, 3.0), dpi=200)
  ax_m3.bar(np.arange(12), m3_qty, width=0.45, color=COLOR_PRIMARY)
  ax_m3.set_xticks(np.arange(12))
  ax_m3.set_xticklabels(months_labels)
  ax_m3.set_yscale("symlog", linthresh=100)
  ax_m3.set_title(f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}", fontweight="bold")
  plt.close(fig3_mpl)

  fig4_mpl, ax_m4 = plt.subplots(figsize=(6, 3.0), dpi=200)
  ax_m4.bar(np.arange(len(fams_x)), deliv_fams, width=0.45, color=bar_colors)
  ax_m4.set_xticks(np.arange(len(fams_x)))
  ax_m4.set_xticklabels(fams_x)
  ax_m4.set_yscale("symlog", linthresh=100)
  ax_m4.set_title("TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5", fontweight="bold")
  plt.close(fig4_mpl)

  # ================= BẢNG SỐ LIỆU TỔNG HỢP & NÚT EXCEL =================
  st.markdown("---")
  pct_orders_m = [((m_comp_orders[i] / m_tot_orders[i]) * 100.0) if m_tot_orders[i] > 0 else 0.0 for i in range(12)]
  pct_qty_m = [((m_comp_qty[i] / (m_comp_qty[i] + m_uncomp_qty[i])) * 100.0) if (m_comp_qty[i] + m_uncomp_qty[i]) > 0 else 0.0 for i in range(12)]
  total_m3 = sum(m3_qty)
  pct_fam_contrib = [((m3_qty[i] / total_m3) * 100.0) if total_m3 > 0 else 0.0 for i in range(12)]
  total_yr = sum(deliv_fams)
  pct_yr_share = [((v / total_yr) * 100.0) if total_yr > 0 else 0.0 for v in deliv_fams]

  df_monthly_summary = pd.DataFrame({"Tháng": months_labels, "Lệnh Hoàn Thành": m_comp_orders, "Lệnh Chưa Xong": m_uncomp_orders, "Tỷ Lệ HT Lệnh (%)": pct_orders_m, "SL Hoàn Thành": [int(v) for v in m_comp_qty], "SL Chưa Xong": [int(v) for v in m_uncomp_qty], "Tỷ Lệ HT SL (%)": pct_qty_m})
  df_plan_summary = pd.DataFrame({"Chỉ Tiêu": ["Tổng Kế Hoạch", "Đã Giao Hoàn Thành", "Còn Lại Chưa Xong"], "Số Lượng": [int(tot_qty_all), int(deliv_qty_all), int(rem_qty_all)], "Tỷ Lệ Cơ Cấu (%)": [100.0, (deliv_qty_all / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0, (rem_qty_all / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0]})
  df_family_summary = pd.DataFrame({"Tháng": months_labels, f"SL Sản Xuất ({sel_fam})": [int(v) for v in m3_qty], "Tỷ Lệ Đóng Góp Tháng (%)": pct_fam_contrib})
  df_year_summary = pd.DataFrame({"Mã / Dòng SP": fams_x, "Tổng SL Hoàn Thành Cả Năm": [int(v) for v in deliv_fams], "Tỷ Lệ Cơ Cấu (%)": pct_yr_share})

  col_hdr, col_btn = st.columns([2.5, 1.2])
  with col_hdr:
    st.markdown(f"### 📑 BẢNG SỐ LIỆU TỔNG HỢP & TỶ LỆ HIỆU CHỈNH - {title_clean.upper()}")
  excel_bytes = generate_print_ready_excel(phan_he_code, title_clean, tu_date, den_date, df_monthly_summary, df_plan_summary, df_family_summary, df_year_summary, fig1_mpl, fig2_mpl, fig3_mpl, fig4_mpl)
  with col_btn:
    st.download_button(label="📥 XUẤT BÁO CÁO EXCEL CHUYÊN NGHIỆP", data=excel_bytes, file_name=f"BaoCao_EMIC_{phan_he_code}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")

  t_col1, t_col2 = st.columns([1.6, 1.0])
  with t_col1:
    st.markdown("##### 1. Tiến Độ Sản Xuất Theo Tháng")
    st.dataframe(df_monthly_summary, column_config={"Tỷ Lệ HT Lệnh (%)": st.column_config.ProgressColumn("Tỷ Lệ HT Lệnh", format="%.1f%%", min_value=0, max_value=100), "Tỷ Lệ HT SL (%)": st.column_config.ProgressColumn("Tỷ Lệ HT SL", format="%.1f%%", min_value=0, max_value=100), "SL Hoàn Thành": st.column_config.NumberColumn("SL Hoàn Thành", format="%d"), "SL Chưa Xong": st.column_config.NumberColumn("SL Chưa Xong", format="%d")}, use_container_width=True, hide_index=True)
  with t_col2:
    st.markdown("##### 2. Tổng Quan Chỉ Tiêu Kế Hoạch")
    st.dataframe(df_plan_summary, column_config={"Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn("Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100), "Số Lượng": st.column_config.NumberColumn("Số Lượng", format="%d")}, use_container_width=True, hide_index=True)

  t_col3, t_col4 = st.columns([1.0, 1.0])
  with t_col3:
    st.markdown(f"##### 3. Chi Tiết Sản Lượng Dòng SP ({sel_fam})")
    st.dataframe(df_family_summary, column_config={"Tỷ Lệ Đóng Góp Tháng (%)": st.column_config.ProgressColumn("Tỷ Lệ Đóng Góp Tháng", format="%.1f%%", min_value=0, max_value=100), f"SL Sản Xuất ({sel_fam})": st.column_config.NumberColumn(f"SL Sản Xuất ({sel_fam})", format="%d")}, use_container_width=True, hide_index=True)
  with t_col4:
    st.markdown("##### 4. Tổng Sản Lượng Cả Năm Các Mã Đầu 5")
    st.dataframe(df_year_summary, column_config={"Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn("Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100), "Tổng SL Hoàn Thành Cả Năm": st.column_config.NumberColumn("Tổng SL Cả Năm", format="%d")}, use_container_width=True, hide_index=True)


# TAB 1 TƯƠNG TỰ (Viết ngắn gọn để tránh quá tải)
with tab_vat_tu:
  if not df_qa32.empty:
    months_labels = [f"T{i}" for i in range(1, 13)]
    ud01_m, ud02_m, ud03_m, uninspected_m = [0]*12, [0]*12, [0]*12, [0]*12
    ft_qty_m, by_inspected_m, by_uninspected_m = [0.0]*12, [0.0]*12, [0.0]*12
    total_ca_block, total_ft_all = 0.0, 0.0
    for _, r in df_qa32.iterrows():
      try: m_idx = datetime.strptime(str(r["ngay_ve_dt"]).split()[0], "%Y-%m-%d").month - 1
      except: m_idx = 0
      st_clean = str(r["xac_nhan_sap"]).strip().upper().replace(" ", "") if pd.notna(r["xac_nhan_sap"]) else ""
      ft_val = float(r["ft_qty"]) if ("ft_qty" in r and pd.notna(r["ft_qty"])) else 0.0
      by_val = float(r["by_sample"]) if ("by_sample" in r and pd.notna(r["by_sample"])) else 0.0
      ca_val = float(r["ca_qty"]) if ("ca_qty" in r and pd.notna(r["ca_qty"])) else 0.0
      ft_qty_m[m_idx] += ft_val; total_ft_all += ft_val; total_ca_block += ca_val
      is_uninspected = ("CHƯA" in st_clean or not st_clean or st_clean in ["NAN", "NONE", "❌CHƯAXN"])
      is_ud02 = any(k in st_clean for k in ["02", "UD2", "ĐẶCNHƯỢNG", "DACNHUONG"])
      is_ud03 = any(k in st_clean for k in ["03", "UD3", "TRẢLẠI", "TRALAI", "TỪCHỐI", "TUCHOI", "KHÔNG", "KHONG"])
      is_ud01 = any(k in st_clean for k in ["01", "UD1", "ĐẠT", "DAT"]) and not (is_ud02 or is_ud03)
      if is_uninspected: uninspected_m[m_idx] += 1; by_uninspected_m[m_idx] += by_val
      else:
        by_inspected_m[m_idx] += by_val
        if is_ud02: ud02_m[m_idx] += 1
        elif is_ud03: ud03_m[m_idx] += 1
        else: ud01_m[m_idx] += 1
    
    col1, col2 = st.columns([2.1, 1.0])
    with col1:
      st.markdown('<div class="chart-card">', unsafe_allow_html=True)
      fig1 = make_subplots(specs=[[{"secondary_y": True}]])
      fig1.add_trace(go.Bar(x=months_labels, y=ud01_m, name="UD 01 (Đạt)", marker_color=COLOR_SUCCESS), secondary_y=False)
      fig1.add_trace(go.Bar(x=months_labels, y=ud02_m, name="UD 02 (Đặc nhượng)", marker_color=COLOR_WARNING), secondary_y=False)
      fig1.add_trace(go.Bar(x=months_labels, y=ud03_m, name="UD 03 (Trả lại)", marker_color=COLOR_DANGER), secondary_y=False)
      total_by = [by_inspected_m[i] + by_uninspected_m[i] for i in range(12)]
      fig1.add_trace(go.Scatter(x=months_labels, y=total_by, name="Số mẫu phải kiểm (BY)", line=dict(color=COLOR_PURPLE, width=2, dash="dash")), secondary_y=True)
      fig1.add_trace(go.Scatter(x=months_labels, y=ft_qty_m, name="Tổng số hàng về (FT)", line=dict(color=COLOR_PRIMARY, width=2)), secondary_y=True)
      fig1.update_layout(title=dict(text="BÁO CÁO SỐ LƯỢNG LỆNH KIỂM & TỔNG VẬT TƯ VỀ / SỐ MẪU KIỂM", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5), barmode="stack", margin=dict(l=30, r=30, t=50, b=90), height=380, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5, font=dict(size=11, family="Arial")))
      fig1.update_yaxes(title_text="← Số Lượng Lệnh", secondary_y=False, showgrid=True, gridcolor="#F1F5F9")
      fig1.update_yaxes(title_text="Vật Tư / Mẫu [Log] →", type="log", secondary_y=True, showgrid=False)
      st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})
      st.markdown("</div>", unsafe_allow_html=True)
    with col2:
      st.markdown('<div class="chart-card">', unsafe_allow_html=True)
      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      pct_ok = (ok_cnt / total_ft_all * 100) if total_ft_all > 0 else 0
      pct_block = (total_ca_block / total_ft_all * 100) if total_ft_all > 0 else 0
      fig2 = go.Figure(data=[go.Pie(labels=["Vật tư Đạt", "Bị Block (Lỗi)"], values=[ok_cnt, total_ca_block], hole=0.6, marker=dict(colors=[COLOR_SUCCESS, COLOR_DANGER], line=dict(color='#FFFFFF', width=3)), text=[f"<b>Vật tư Đạt</b><br>{pct_ok:.1f}%<br>({ok_cnt:,.0f})", f"<b>Bị Block (Lỗi)</b><br>{pct_block:.1f}%<br>({total_ca_block:,.0f})"], textinfo="text", textposition="outside", textfont=dict(size=12, family="Arial", color=[COLOR_SUCCESS, COLOR_DANGER]), direction="clockwise", sort=False)])
      fig2.update_layout(title=dict(text="TỶ LỆ VẬT TƯ ĐẠT VS BỊ BLOCK LỖI", font=dict(size=14, color=COLOR_TEXT, family="Arial", weight="bold"), x=0.5), margin=dict(l=30, r=30, t=50, b=90), height=380, paper_bgcolor="#FFFFFF", showlegend=False, annotations=[dict(text=f"<b>TỔNG VẬT TƯ VỀ</b><br><span style='font-size:16px'>{int(total_ft_all):,}</span>", x=0.5, y=0.5, font_size=12, font_family="Arial", showarrow=False)])
      st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})
      st.markdown("</div>", unsafe_allow_html=True)

with tab_co_khi: render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
with tab_tuti: render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")
with tab_cong_to: render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")