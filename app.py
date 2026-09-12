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
import streamlit as st

# ================= 1. CẤU HÌNH DASHBOARD & BỘ STYLE CSS CHUẨN DESKTOP =================
st.set_page_config(
    page_title="EMIC QC Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        header[data-testid="stHeader"] { display: none !important; }
        
        .main .block-container, div[data-testid="stAppViewBlockContainer"] {
            padding-top: 0.3rem !important;
            padding-bottom: 0.3rem !important;
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
            max-width: 100% !important;
        }

        .stApp {
            background-color: #F8FAFC;
            font-family: system-ui, -apple-system, sans-serif;
        }

        /* Styling Tabs Navigation Căn Giữa Đỉnh Trang */
        .stTabs [data-baseweb="tab-list"] {
            justify-content: center !important;
            gap: 8px !important;
            background-color: transparent !important;
            padding: 2px 0px 8px 0px !important;
            border-bottom: 1px solid #E2E8F0 !important;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #94A3B8 !important;
            color: #FFFFFF !important;
            border-radius: 6px !important;
            padding: 5px 18px !important;
            border: none !important;
            font-weight: 600 !important;
            font-size: 13px !important;
        }
        .stTabs [aria-selected="true"] {
            background-color: #3B82F6 !important;
            color: #FFFFFF !important;
            box-shadow: 0 2px 4px rgba(59, 130, 246, 0.3) !important;
        }
        
        /* Card Khung Bao Cho Từng Biểu Đồ (Khắc Phục Ô Trắng Thừa) */
        .chart-card {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            padding: 8px 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            margin-bottom: 8px;
        }

        div[data-testid="stVerticalBlock"] > div {
            gap: 0.2rem !important;
        }
        
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 1px solid #E2E8F0;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Cấu hình Matplotlib Sắc Nét HD
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [
    "DejaVu Sans",
    "Liberation Sans",
    "Arial",
    "sans-serif",
]
plt.rcParams["font.size"] = 8
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8

# Bảng Màu Chuẩn Nguyên Bản 100%
COLOR_SUCCESS = "#10B981"  # Xanh lá (UD01 / SL Hoàn thành / Đạt)
COLOR_PRIMARY = "#3B82F6"  # Xanh dương (Lệnh Hoàn Thành / FT)
COLOR_DANGER = "#EF4444"  # Đỏ (UD03 / Block / Lệnh chưa xong)
COLOR_WARNING = "#F59E0B"  # Cam (UD02 / SL chưa xong)
COLOR_PURPLE = "#A855F7"  # Tím (BY Sample / PT)
COLOR_CYAN = "#06B6D4"  # Cyan (VT)

DISTINCT_COLORS = [
    "#3B82F6",
    "#10B981",
    "#F59E0B",
    "#EF4444",
    "#A855F7",
    "#06B6D4",
    "#F43F5E",
    "#84CC16",
    "#F97316",
    "#14B8A6",
]


class Theme:
  SURFACE = "#FFFFFF"
  TEXT_PRIMARY = "#0F172A"


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


def log_formatter(x, pos):
  if x <= 0:
    return "0"
  if x >= 1e6:
    return f"{x*1e-6:.1f}M"
  if x >= 1e3:
    return f"{x*1e-3:.0f}K"
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
      conn,
      params=(tu_iso, den_iso),
  )
  df_coois = pd.read_sql_query(
      "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt <="
      " ?",
      conn,
      params=(tu_iso, den_iso),
  )
  conn.close()
  return df_qa32, df_coois


# ================= 2. HÀM TẠO FILE EXCEL CHUẨN IN ẤN TRÌNH BÀY CHUYÊN NGHIỆP =================
def generate_print_ready_excel(
    phan_he_code,
    title_clean,
    tu_date,
    den_date,
    df_monthly,
    df_plan,
    df_family,
    df_year,
    fig1_mpl,
    fig2_mpl,
    fig3_mpl,
    fig4_mpl,
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

  fill_header = PatternFill(
      start_color="1F4E79", end_color="1F4E79", fill_type="solid"
  )
  fill_zebra = PatternFill(
      start_color="F9FBFD", end_color="F9FBFD", fill_type="solid"
  )

  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )
  header_border = Border(
      left=Side(style="thin", color="FFFFFF"),
      right=Side(style="thin", color="FFFFFF"),
      top=Side(style="medium", color="1F4E79"),
      bottom=Side(style="medium", color="1F4E79"),
  )

  align_center = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

  sheets_data = [
      ("TienDo_Thang", "1. TIẾN ĐỘ SẢN XUẤT THEO THÁNG", df_monthly),
      ("TongQuan_KeHoach", "2. TỔNG QUAN CHỈ TIÊU KẾ HOẠCH", df_plan),
      ("Dong_SanPham", "3. CHI TIẾT THEO DÒNG SẢN PHẨM", df_family),
      ("SanLuong_CaNam", "4. TỔNG SẢN LƯỢNG CẢ NĂM MÃ ĐẦU 5", df_year),
  ]

  for sheet_name, section_title, df_table in sheets_data:
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws["A1"] = "TỔNG CÔNG TY THIẾT BỊ ĐIỆN EMIC - PHÒNG QUẢN LÝ CHẤT LƯỢNG (QC)"
    ws["A1"].font = font_company

    ws["A2"] = f"BÁO CÁO TỔNG HỢP SẢN XUẤT - {title_clean.upper()}"
    ws["A2"].font = font_title

    ws["A3"] = (
        f"Giai đoạn: Từ {tu_date.strftime('%d/%m/%Y')} đến"
        f" {den_date.strftime('%d/%m/%Y')} | Ngày xuất:"
        f" {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
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
      is_even = row_idx % 2 == 0
      row_fill = fill_zebra if is_even else PatternFill(fill_type=None)

      for col_idx, val in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx)
        col_name_lower = str(df_table.columns[col_idx - 1]).lower()

        if isinstance(val, (int, float, np.number)):
          if "%" in col_name_lower or "tỷ lệ" in col_name_lower:
            cell.value = float(val) / 100.0 if val > 1.0 else float(val)
            cell.number_format = "0.0%"
          else:
            cell.value = float(val)
            cell.number_format = (
                "#,##0" if float(val).is_integer() else "#,##0.00"
            )
          cell.alignment = align_right
        else:
          cell.value = str(val)
          cell.alignment = (
              align_center
              if col_idx == 1 or len(str(val)) < 10
              else align_left
          )

        cell.font = font_data
        if row_fill.fill_type:
          cell.fill = row_fill
        cell.border = thin_border

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        if cell.row >= start_row and cell.value is not None:
          max_len = max(max_len, len(str(cell.value)))
      ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

  # Sheet Đính Kèm Hình Ảnh Biểu Đồ
  ws_img = wb.create_sheet(title="BieuDo_Visual")
  ws_img.views.sheetView[0].showGridLines = True
  ws_img["A1"] = "HÌNH ẢNH CÁC BIỂU ĐỒ BÁO CÁO CÂN ĐỐI"
  ws_img["A1"].font = font_title

  mpl_figs = [
      (fig1_mpl, "B3"),
      (fig2_mpl, "K3"),
      (fig3_mpl, "B22"),
      (fig4_mpl, "K22"),
  ]
  for f_mpl, cell_pos in mpl_figs:
    if f_mpl is not None:
      buf = io.BytesIO()
      f_mpl.savefig(
          buf, format="png", dpi=200, bbox_inches="tight", facecolor="#FFFFFF"
      )
      buf.seek(0)
      img = OpenpyxlImage(buf)
      ws_img.add_image(img, cell_pos)

  wb.save(output)
  return output.getvalue()


# ================= 3. SIDEBAR BỘ LỌC NGÀY =================
st.sidebar.title("🎛️ BỘ LỌC DỮ LIỆU")
col_tu, col_den = st.sidebar.columns(2)
with col_tu:
  tu_date = st.date_input("Từ ngày", date(datetime.now().year, 1, 1))
with col_den:
  den_date = st.date_input("Đến ngày", date.today())

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Cập Nhật Dữ Liệu", use_container_width=True):
  st.cache_data.clear()
  st.rerun()

df_qa32, df_coois = load_data(tu_date, den_date)

# ================= 4. NAVIGATION TABS CĂN GIỮA ĐỈNH TRANG =================
tab_vat_tu, tab_co_khi, tab_tuti, tab_cong_to = st.tabs([
    "📋 Báo Cáo Vật Tư",
    "⚙️ Báo Cáo Cơ Khí",
    "🔌 Báo Cáo TU/TI",
    "⚡ Báo Cáo Công Tơ",
])

# ================= 5. TAB 1: BÁO CÁO VẬT TƯ =================
with tab_vat_tu:
  if df_qa32.empty:
    st.info("💡 Chưa có dữ liệu QA32 trong khoảng thời gian đã chọn.")
  else:
    months_labels = [f"T{i}" for i in range(1, 13)]
    ud01_m, ud02_m, ud03_m, uninspected_m = [0] * 12, [0] * 12, [0] * 12, [0] * 12
    ft_qty_m, by_inspected_m, by_uninspected_m = (
        [0.0] * 12,
        [0.0] * 12,
        [0.0] * 12,
    )
    total_ca_block, total_ft_all = 0.0, 0.0
    top_block_dict = {}

    for _, r in df_qa32.iterrows():
      try:
        m_idx = (
            datetime.strptime(str(r["ngay_ve_dt"]).split()[0], "%Y-%m-%d").month
            - 1
        )
      except:
        m_idx = 0
      if not (0 <= m_idx < 12):
        m_idx = 0

      st_clean = (
          str(r["xac_nhan_sap"]).strip().upper().replace(" ", "")
          if pd.notna(r["xac_nhan_sap"])
          else ""
      )
      ft_val = (
          float(r["ft_qty"])
          if ("ft_qty" in r and pd.notna(r["ft_qty"]))
          else 0.0
      )
      by_val = (
          float(r["by_sample"])
          if ("by_sample" in r and pd.notna(r["by_sample"]))
          else 0.0
      )
      ca_val = (
          float(r["ca_qty"])
          if ("ca_qty" in r and pd.notna(r["ca_qty"]))
          else 0.0
      )

      ft_qty_m[m_idx] += ft_val
      total_ft_all += ft_val
      total_ca_block += ca_val

      is_uninspected = (
          "CHƯA" in st_clean
          or not st_clean
          or st_clean in ["NAN", "NONE", "❌CHƯAXN"]
      )
      is_ud02 = any(
          k in st_clean for k in ["02", "UD2", "ĐẶCNHƯỢNG", "DACNHUONG"]
      )
      is_ud03 = any(
          k in st_clean
          for k in [
              "03",
              "UD3",
              "TRẢLẠI",
              "TRALAI",
              "TỪCHỐI",
              "TUCHOI",
              "KHÔNG",
              "KHONG",
          ]
      )
      is_ud01 = any(k in st_clean for k in ["01", "UD1", "ĐẠT", "DAT"]) and not (
          is_ud02 or is_ud03
      )

      if is_uninspected:
        uninspected_m[m_idx] += 1
        by_uninspected_m[m_idx] += by_val
      else:
        by_inspected_m[m_idx] += by_val
        if is_ud02:
          ud02_m[m_idx] += 1
        elif is_ud03:
          ud03_m[m_idx] += 1
        else:
          ud01_m[m_idx] += 1

      ma_vt_str = str(r["ma_vt"]).strip() if pd.notna(r["ma_vt"]) else ""
      ten_vt_str = str(r["ten_vt"]).strip() if pd.notna(r["ten_vt"]) else ""
      ncc_str = str(r["ncc"]).strip() if pd.notna(r["ncc"]) else ""

      if (
          is_ud02
          or is_ud03
          or ca_val > 0
          or (st_clean and not is_ud01 and not is_uninspected)
      ):
        key = (ma_vt_str, ncc_str)
        if key not in top_block_dict:
          top_block_dict[key] = {
              "ma_vt": ma_vt_str,
              "ten_vt": ten_vt_str,
              "ncc": ncc_str,
              "ud02": 0,
              "ud03": 0,
              "ca_block": 0.0,
              "ft_total": 0.0,
          }
        if is_ud02:
          top_block_dict[key]["ud02"] += 1
        if is_ud03:
          top_block_dict[key]["ud03"] += 1
        top_block_dict[key]["ca_block"] += ca_val
        top_block_dict[key]["ft_total"] += ft_val

    col1, col2 = st.columns([2.1, 1.0])

    with col1:
      st.markdown('<div class="chart-card">', unsafe_allow_html=True)
      fig1 = plt.figure(figsize=(8.2, 3.2), dpi=200)
      fig1.patch.set_facecolor(Theme.SURFACE)

      ax1 = fig1.add_subplot(111)
      ax1.set_facecolor(Theme.SURFACE)
      ax2 = ax1.twinx()

      ax1.set_zorder(1)
      ax2.set_zorder(2)
      ax1.patch.set_alpha(0.0)

      for spine in ["top"]:
        ax1.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)

      x = np.arange(12)
      w = 0.38

      b_ud01 = ax1.bar(x, ud01_m, width=w, color=COLOR_SUCCESS, label="UD 01 (Đạt)")
      b_ud02 = ax1.bar(
          x,
          ud02_m,
          width=w,
          bottom=ud01_m,
          color=COLOR_WARNING,
          label="UD 02 (Đặc nhượng)",
      )
      bot_03 = [ud01_m[i] + ud02_m[i] for i in range(12)]
      b_ud03 = ax1.bar(
          x,
          ud03_m,
          width=w,
          bottom=bot_03,
          color=COLOR_DANGER,
          label="UD 03 (Trả lại)",
      )

      max_order_val = max(
          [ud01_m[i] + ud02_m[i] + ud03_m[i] for i in range(12)] + [1]
      )
      ax1.set_ylim(0, max_order_val * 1.25)

      for i in range(12):
        un_cnt = uninspected_m[i]
        if un_cnt > 0:
          ax1.text(
              x[i],
              (ud01_m[i] + ud02_m[i] + ud03_m[i]) + (max_order_val * 0.02),
              f"{un_cnt}",
              ha="center",
              va="bottom",
              fontweight="bold",
              fontsize=7,
              color=COLOR_PURPLE,
          )

      total_by_sample_m = [
          by_inspected_m[i] + by_uninspected_m[i] for i in range(12)
      ]
      line_ft = ax2.plot(
          x,
          ft_qty_m,
          color=COLOR_PRIMARY,
          marker="o",
          linewidth=1.8,
          label="Tổng số hàng về (FT)",
      )
      line_by = ax2.plot(
          x,
          total_by_sample_m,
          color=COLOR_PURPLE,
          marker="s",
          linewidth=1.8,
          linestyle="--",
          label="Số mẫu phải kiểm (BY)",
      )

      ax2.set_yscale("symlog", linthresh=100)
      ax2.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

      ax1.set_xticks(x)
      ax1.set_xticklabels(months_labels, fontweight="bold", fontsize=8)
      ax1.set_xlim(-0.6, 11.6)
      ax1.set_ylabel(
          "← Số Lượng Lệnh (Trục Trái)",
          fontweight="bold",
          color=COLOR_SUCCESS,
          fontsize=8,
      )
      ax2.set_ylabel(
          "Số Lượng Vật Tư / Mẫu [Log] (Trục Phải) →",
          fontweight="bold",
          color=COLOR_PRIMARY,
          fontsize=8,
      )
      ax1.set_title(
          "BÁO CÁO SỐ LƯỢNG LỆNH KIỂM & TỔNG VẬT TƯ VỀ / SỐ MẪU KIỂM",
          fontweight="bold",
          fontsize=9.5,
          color=Theme.TEXT_PRIMARY,
          pad=10,
      )

      all_handles = [b_ud01, b_ud02, b_ud03, line_by[0], line_ft[0]]
      all_labels = [
          "UD 01 (Đạt)",
          "UD 02 (Đặc nhượng)",
          "UD 03 (Trả lại)",
          "Số mẫu phải kiểm (BY)",
          "Tổng số hàng về (FT)",
      ]
      ax1.legend(
          all_handles,
          all_labels,
          loc="upper center",
          bbox_to_anchor=(0.5, -0.18),
          frameon=False,
          fontsize=7.5,
          ncol=3,
      )

      fig1.subplots_adjust(
          top=0.88, bottom=0.20, left=0.08, right=0.92, wspace=0.15
      )
      st.pyplot(fig1, use_container_width=True)
      st.markdown("</div>", unsafe_allow_html=True)

    with col2:
      st.markdown('<div class="chart-card">', unsafe_allow_html=True)
      fig_pie = plt.figure(figsize=(3.8, 3.2), dpi=200)
      fig_pie.patch.set_facecolor(Theme.SURFACE)
      ax_pie = fig_pie.add_subplot(111)
      ax_pie.set_facecolor(Theme.SURFACE)

      ax_pie.set_aspect("equal")

      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      pct_ok = (ok_cnt / total_ft_all * 100) if total_ft_all > 0 else 0
      pct_block = (
          (total_ca_block / total_ft_all * 100) if total_ft_all > 0 else 0
      )

      if total_ft_all > 0:
        wedges, texts = ax_pie.pie(
            [ok_cnt, total_ca_block],
            labels=[
                f"Vật tư Đạt\n{ok_cnt:,.0f}\n{pct_ok:.1f}%",
                f"Bị Block (Lỗi)\n{total_ca_block:,.0f}\n{pct_block:.1f}%",
            ],
            colors=[COLOR_SUCCESS, COLOR_DANGER],
            startangle=140,
            pctdistance=0.6,
            labeldistance=1.18,
            radius=0.82,
            wedgeprops=dict(width=0.35, edgecolor="white", linewidth=2),
        )
        texts[0].set_color(COLOR_SUCCESS)
        texts[0].set_fontweight("bold")
        texts[0].set_fontsize(7.5)
        if len(texts) > 1:
          texts[1].set_color(COLOR_DANGER)
          texts[1].set_fontweight("bold")
          texts[1].set_fontsize(7.5)

        ax_pie.text(
            0,
            0,
            f"TỔNG VẬT TƯ VỀ\n{total_ft_all:,.0f}",
            ha="center",
            va="center",
            fontweight="bold",
            fontsize=8,
            color=Theme.TEXT_PRIMARY,
        )
      else:
        ax_pie.text(0, 0, "Chưa có dữ liệu", ha="center", fontsize=8)
        ax_pie.axis("off")

      ax_pie.set_title(
          "TỶ LỆ VẬT TƯ ĐẠT VS BỊ BLOCK LỖI",
          fontweight="bold",
          fontsize=9.5,
          color=Theme.TEXT_PRIMARY,
          pad=10,
      )
      fig_pie.subplots_adjust(top=0.88, bottom=0.08, left=0.06, right=0.94)
      st.pyplot(fig_pie, use_container_width=True)
      st.markdown("</div>", unsafe_allow_html=True)

    sorted_blocks = sorted(
        top_block_dict.values(),
        key=lambda x: (x["ud03"] + x["ud02"], x["ca_block"], x["ft_total"]),
        reverse=True,
    )
    if sorted_blocks:
      df_block = pd.DataFrame(sorted_blocks)
      df_block["Tổng SL Block (CA) / SL Về"] = df_block.apply(
          lambda r: f"{r['ca_block']:,.0f} / {r['ft_total']:,.0f}", axis=1
      )
      df_block = df_block[[
          "ma_vt",
          "ten_vt",
          "ncc",
          "ud02",
          "ud03",
          "Tổng SL Block (CA) / SL Về",
      ]]
      df_block.columns = [
          "Mã Vật Tư",
          "Tên Vật Tư",
          "Nhà Cung Cấp",
          "Số Lượt UD 02",
          "Số Lượt UD 03",
          "Tổng SL Block (CA) / SL Về",
      ]
      st.dataframe(df_block, use_container_width=True, hide_index=True)


# ================= 6. HÀM CHUNG BÁO CÁO COOIS (QUY CHUẨN CHIỀU CAO BẰNG CHẰN CHẶN) =================
def render_coois_tab_layout(phan_he_code, title_text):
  df_sub = (
      df_coois[df_coois["phan_he"] == phan_he_code]
      if not df_coois.empty
      else pd.DataFrame()
  )

  if df_sub.empty:
    st.info(f"💡 Chưa có dữ liệu sản xuất cho phân hệ {title_text}.")
    return

  months_labels = [f"T{i}" for i in range(1, 13)]
  m_comp_qty, m_uncomp_qty = [0.0] * 12, [0.0] * 12
  m_tot_orders, m_uncomp_orders = [0] * 12, [0] * 12
  tot_qty_all, deliv_qty_all = 0.0, 0.0

  for _, r in df_sub.iterrows():
    try:
      m_idx = (
          datetime.strptime(str(r["ngay_lenh_dt"]).split()[0], "%Y-%m-%d").month
          - 1
      )
    except:
      m_idx = 0
    if not (0 <= m_idx < 12):
      m_idx = 0

    sl_t, sl_h = float(r["sl_tong"]), float(r["sl_ht"])
    uncomp_q = max(0.0, sl_t - sl_h)

    tot_qty_all += sl_t
    deliv_qty_all += sl_h
    m_comp_qty[m_idx] += sl_h
    m_uncomp_qty[m_idx] += uncomp_q

    m_tot_orders[m_idx] += 1
    if sl_h < sl_t:
      m_uncomp_orders[m_idx] += 1

  title_clean = clean_emoji(title_text)

  # HÀNG 1: TIẾN ĐỘ SẢN XUẤT (2.1) & DONUT CHART (1.0) - CÙNG FIGSIZE (..., 3.1)
  col1, col2 = st.columns([2.1, 1.0])

  with col1:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    fig1, ax1 = plt.subplots(figsize=(7.8, 3.1), dpi=200)
    fig1.patch.set_facecolor(Theme.SURFACE)
    ax1.set_facecolor(Theme.SURFACE)
    ax2 = ax1.twinx()

    for spine in ["top"]:
      ax1.spines[spine].set_visible(False)
      ax2.spines[spine].set_visible(False)

    x = np.arange(12)
    w = 0.35
    m_comp_orders = [m_tot_orders[i] - m_uncomp_orders[i] for i in range(12)]

    ax1.bar(
        x - w / 2,
        m_comp_orders,
        width=w,
        color=COLOR_PRIMARY,
        label="Lệnh Hoàn Thành",
    )
    ax1.bar(
        x - w / 2,
        m_uncomp_orders,
        width=w,
        bottom=m_comp_orders,
        color=COLOR_DANGER,
        label="Lệnh Chưa Xong",
    )

    ax2.bar(
        x + w / 2,
        m_comp_qty,
        width=w,
        color=COLOR_SUCCESS,
        label="SL Hoàn Thành",
    )
    ax2.bar(
        x + w / 2,
        m_uncomp_qty,
        width=w,
        bottom=m_comp_qty,
        color=COLOR_WARNING,
        label="SL Chưa Xong",
    )

    ax2.set_yscale("symlog", linthresh=100)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    max_q = max([m_comp_qty[i] + m_uncomp_qty[i] for i in range(12)] + [100])
    ax2.set_ylim(0, max_q * 3.2)

    for i in range(12):
      t_qty = m_comp_qty[i] + m_uncomp_qty[i]
      pct = (m_comp_qty[i] / t_qty * 100) if t_qty > 0 else 0
      if t_qty > 0:
        ax2.text(
            x[i] + w / 2,
            t_qty * 1.12,
            f"{pct:.0f}%",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=7,
            color=COLOR_SUCCESS,
        )

    ax1.set_title(
        f"TIẾN ĐỘ SẢN XUẤT - {title_clean}",
        fontweight="bold",
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )
    ax1.set_ylabel(
        "← Tổng Lệnh (Trục Trái)",
        fontweight="bold",
        color=COLOR_PRIMARY,
        fontsize=8,
    )
    ax2.set_ylabel(
        "Số Lượng Giao [Log] (Trục Phải) →",
        fontweight="bold",
        color=COLOR_SUCCESS,
        fontsize=8,
    )
    ax1.set_xticks(x)
    ax1.set_xticklabels(months_labels, fontweight="bold", fontsize=8)
    ax1.set_xlim(-0.6, 11.6)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        frameon=False,
        fontsize=7.5,
        ncol=4,
    )

    fig1.subplots_adjust(top=0.88, bottom=0.20, left=0.09, right=0.91)
    st.pyplot(fig1, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

  with col2:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    fig2, ax3 = plt.subplots(figsize=(3.8, 3.1), dpi=200)
    fig2.patch.set_facecolor(Theme.SURFACE)
    ax3.set_facecolor(Theme.SURFACE)

    # KHÓA TỶ LỆ TRÒN NATIVE 1:1 CHỐNG MÉO
    ax3.set_aspect("equal")

    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    if tot_qty_all > 0:
      wedges, texts = ax3.pie(
          [deliv_qty_all, rem_qty_all],
          labels=[
              f"Hoàn thành\n{pct_deliv:.1f}%\n({int(deliv_qty_all):,})",
              f"Chưa xong\n{pct_rem:.1f}%\n({int(rem_qty_all):,})",
          ],
          colors=[COLOR_SUCCESS, COLOR_WARNING],
          startangle=140,
          pctdistance=0.6,
          labeldistance=1.18,
          radius=0.85,
          wedgeprops=dict(width=0.35, edgecolor="white", linewidth=2),
      )
      texts[0].set_color(COLOR_SUCCESS)
      texts[0].set_fontweight("bold")
      texts[0].set_fontsize(7.5)
      if len(texts) > 1:
        texts[1].set_color(COLOR_DANGER)
        texts[1].set_fontweight("bold")
        texts[1].set_fontsize(7.5)

      ax3.text(
          0,
          0,
          f"TỔNG KẾ HOẠCH\n{int(tot_qty_all):,}",
          ha="center",
          va="center",
          fontweight="bold",
          fontsize=8,
          color=Theme.TEXT_PRIMARY,
      )
    else:
      ax3.text(0, 0, "Chưa có dữ liệu", ha="center", fontsize=8)
      ax3.axis("off")

    ax3.set_title(
        "TỶ LỆ HOÀN THÀNH TỔNG QUAN",
        fontweight="bold",
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )
    fig2.subplots_adjust(top=0.88, bottom=0.08, left=0.06, right=0.94)
    st.pyplot(fig2, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

  # HÀNG 2: BỘ LỌC ĐẶT NGOÀI & 2 ĐỒ THỊ BÊN DƯỚI THẲNG HÀNG 100%
  sub_5 = (
      df_sub[
          df_sub["ma_tp"]
          .astype(str)
          .str.split(".")
          .str[0]
          .str.lstrip("0")
          .str.startswith("5")
      ].copy()
      if not df_sub.empty
      else pd.DataFrame()
  )

  raw_fams = (
      [
          str(x).strip()
          for x in sub_5["mat_prefix"].unique()
          if pd.notna(x)
          and str(x).strip()
          and str(x).strip().lower() not in ["none", "nan"]
      ]
      if not sub_5.empty
      else []
  )
  clean_fams = sorted(list(set(raw_fams)))
  available_fams = ["Tất cả dòng sản phẩm"] + clean_fams

  # Ô Lọc Dòng SP Đặt Trực Tiếp (Triệt Tiêu Ô Trắng Thừa)
  sel_fam = st.selectbox(
      "🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}"
  )

  col3, col4 = st.columns([1, 1])

  # --- DƯỚI TRÁI: BIỂU ĐỒ SẢN LƯỢNG DÒNG SP ---
  with col3:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    m3_qty = [0.0] * 12
    if not sub_5.empty:
      sub_5_df = sub_5.copy()
      sub_5_df["month"] = pd.to_datetime(
          sub_5_df["ngay_lenh_dt"], errors="coerce"
      ).dt.month
      sub_filtered = (
          sub_5_df[sub_5_df["mat_prefix"] == sel_fam]
          if sel_fam != "Tất cả dòng sản phẩm"
          else sub_5_df
      )
      for _, r in sub_filtered.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          m3_qty[int(m_val) - 1] += float(r["sl_ht"])

    defect_rate_m = [0.0] * 12

    fig3, ax_b3 = plt.subplots(figsize=(6.0, 2.7), dpi=200)
    fig3.patch.set_facecolor(Theme.SURFACE)
    ax_b3.set_facecolor(Theme.SURFACE)

    for spine in ["top"]:
      ax_b3.spines[spine].set_visible(False)
    ax_b3_right = ax_b3.twinx()
    for spine in ["top"]:
      ax_b3_right.spines[spine].set_visible(False)

    ax_b3.bar(
        x, m3_qty, width=0.42, color=COLOR_PRIMARY, alpha=0.9, label="SL Sản Xuất"
    )
    ax_b3.set_yscale("symlog", linthresh=100)
    ax_b3.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    max_v_b3 = max(m3_qty + [100])
    ax_b3.set_ylim(0, max_v_b3 * 3.2)

    for i in range(12):
      v = m3_qty[i]
      if v > 0:
        ax_b3.text(
            x[i],
            v * 1.18,
            f"{int(v):,}",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=7,
            color=COLOR_PRIMARY,
        )

    ax_b3_right.plot(
        x,
        defect_rate_m,
        color=COLOR_DANGER,
        marker="o",
        linewidth=1.2,
        label="Tỷ lệ sai hỏng (%)",
    )

    ax_b3.set_xticks(x)
    ax_b3.set_xticklabels(months_labels, fontweight="bold", fontsize=8)
    ax_b3.set_xlim(-0.6, 11.6)
    ax_b3.set_ylabel(
        "SL Hoàn Thành [Log]", fontweight="bold", color=COLOR_PRIMARY, fontsize=8
    )
    ax_b3_right.set_ylabel(
        "Sai Hỏng (%)", fontweight="bold", color=COLOR_DANGER, fontsize=8
    )
    ax_b3_right.set_ylim(-1.0, 5.0)
    ax_b3.set_title(
        f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}",
        fontweight="bold",
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )

    fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.12, right=0.88)
    st.pyplot(fig3, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

  # --- DƯỚI PHẢI: TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5 ---
  with col4:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    fig4, ax_b4 = plt.subplots(figsize=(6.0, 2.7), dpi=200)
    fig4.patch.set_facecolor(Theme.SURFACE)
    ax_b4.set_facecolor(Theme.SURFACE)

    for spine in ["top"]:
      ax_b4.spines[spine].set_visible(False)
    ax_b4_right = ax_b4.twinx()
    for spine in ["top"]:
      ax_b4_right.spines[spine].set_visible(False)

    if not sub_5.empty:
      summary_fams = (
          sub_5.groupby("mat_prefix")[["sl_ht"]].sum().reset_index()
      )
      fams_x = [
          str(val)
          for val in summary_fams["mat_prefix"].tolist()
          if pd.notna(val)
          and str(val).strip()
          and str(val).strip().lower() not in ["none", "nan"]
      ]
      if not fams_x:
        fams_x, deliv_fams, defect_fams = ["Trống"], [0.0], [0.0]
      else:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x = summary_fams["mat_prefix"].tolist()
        deliv_fams = summary_fams["sl_ht"].values
        defect_fams = [0.0] * len(fams_x)
    else:
      fams_x, deliv_fams, defect_fams = ["Không có SP"], [0.0], [0.0]

    x_b4 = np.arange(len(fams_x))
    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    bars_deliv = ax_b4.bar(
        x_b4,
        deliv_fams,
        width=0.45,
        label="SL Hoàn Thành Cả Năm",
        color=bar_colors,
    )
    ax_b4.set_yscale("symlog", linthresh=100)
    ax_b4.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    max_v_b4 = max(list(deliv_fams) + [100])
    ax_b4.set_ylim(0, max_v_b4 * 3.5)

    for i in range(len(fams_x)):
      v = deliv_fams[i]
      if v > 0:
        ax_b4.text(
            x_b4[i],
            v * 1.18,
            f"{int(v):,}",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=7.5,
            color=bar_colors[i % len(bar_colors)],
        )

    ax_b4_right.plot(
        x_b4,
        defect_fams,
        color=COLOR_DANGER,
        marker="s",
        linewidth=1.2,
        label="Tỷ lệ sai hỏng (%)",
    )

    ax_b4.set_xticks(x_b4)
    ax_b4.set_xticklabels(fams_x, fontweight="bold", fontsize=8)
    ax_b4.set_ylabel(
        "Số Lượng SP [Log]", fontweight="bold", color=COLOR_SUCCESS, fontsize=8
    )
    ax_b4_right.set_ylabel(
        "Sai Hỏng (%)", fontweight="bold", color=COLOR_DANGER, fontsize=8
    )
    ax_b4_right.set_ylim(-1.0, 5.0)
    ax_b4.set_title(
        "TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5",
        fontweight="bold",
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )

    fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.12, right=0.88)
    st.pyplot(fig4, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

  # ================= 7. BẢNG SỐ LIỆU TỔNG HỢP & NÚT XUẤT EXCEL IN A4 =================
  st.markdown("---")

  pct_orders_m = [
      ((m_comp_orders[i] / m_tot_orders[i]) * 100.0)
      if m_tot_orders[i] > 0
      else 0.0
      for i in range(12)
  ]
  pct_qty_m = [
      ((m_comp_qty[i] / (m_comp_qty[i] + m_uncomp_qty[i])) * 100.0)
      if (m_comp_qty[i] + m_uncomp_qty[i]) > 0
      else 0.0
      for i in range(12)
  ]

  total_m3 = sum(m3_qty)
  pct_fam_contrib = [
      ((m3_qty[i] / total_m3) * 100.0) if total_m3 > 0 else 0.0
      for i in range(12)
  ]

  total_yr = sum(deliv_fams)
  pct_yr_share = [
      ((v / total_yr) * 100.0) if total_yr > 0 else 0.0 for v in deliv_fams
  ]

  df_monthly_summary = pd.DataFrame({
      "Tháng": months_labels,
      "Lệnh Hoàn Thành": m_comp_orders,
      "Lệnh Chưa Xong": m_uncomp_orders,
      "Tỷ Lệ HT Lệnh (%)": pct_orders_m,
      "SL Hoàn Thành": [int(v) for v in m_comp_qty],
      "SL Chưa Xong": [int(v) for v in m_uncomp_qty],
      "Tỷ Lệ HT SL (%)": pct_qty_m,
  })

  df_plan_summary = pd.DataFrame({
      "Chỉ Tiêu": ["Tổng Kế Hoạch", "Đã Giao Hoàn Thành", "Còn Lại Chưa Xong"],
      "Số Lượng": [int(tot_qty_all), int(deliv_qty_all), int(rem_qty_all)],
      "Tỷ Lệ Cơ Cấu (%)": [
          100.0,
          (deliv_qty_all / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0,
          (rem_qty_all / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0,
      ],
  })

  df_family_summary = pd.DataFrame({
      "Tháng": months_labels,
      f"SL Sản Xuất ({sel_fam})": [int(v) for v in m3_qty],
      "Tỷ Lệ Đóng Góp Tháng (%)": pct_fam_contrib,
  })

  df_year_summary = pd.DataFrame({
      "Mã / Dòng SP": fams_x,
      "Tổng SL Hoàn Thành Cả Năm": [int(v) for v in deliv_fams],
      "Tỷ Lệ Cơ Cấu (%)": pct_yr_share,
  })

  col_hdr, col_btn = st.columns([2.5, 1.2])
  with col_hdr:
    st.markdown(
        f"### 📑 BẢNG SỐ LIỆU TỔNG HỢP & TỶ LỆ HIỆU CHỈNH -"
        f" {title_clean.upper()}"
    )

  excel_bytes = generate_print_ready_excel(
      phan_he_code,
      title_clean,
      tu_date,
      den_date,
      df_monthly_summary,
      df_plan_summary,
      df_family_summary,
      df_year_summary,
      fig1,
      fig2,
      fig3,
      fig4,
  )

  with col_btn:
    st.download_button(
        label="📥 XUẤT BÁO CÁO EXCEL (CHUẨN IN A4 & KÈM BIỂU ĐỒ)",
        data=excel_bytes,
        file_name=f"BaoCao_EMIC_{phan_he_code}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

  # Hiển Thị 4 Bảng Số Liệu Dạng Progress Bar Trực Quan
  t_col1, t_col2 = st.columns([1.6, 1.0])
  with t_col1:
    st.markdown("##### 1. Tiến Độ Sản Xuất Theo Tháng")
    st.dataframe(
        df_monthly_summary,
        column_config={
            "Tỷ Lệ HT Lệnh (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ HT Lệnh", format="%.1f%%", min_value=0, max_value=100
            ),
            "Tỷ Lệ HT SL (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ HT SL", format="%.1f%%", min_value=0, max_value=100
            ),
            "SL Hoàn Thành": st.column_config.NumberColumn(
                "SL Hoàn Thành", format="%d"
            ),
            "SL Chưa Xong": st.column_config.NumberColumn(
                "SL Chưa Xong", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )

  with t_col2:
    st.markdown("##### 2. Tổng Quan Chỉ Tiêu Kế Hoạch")
    st.dataframe(
        df_plan_summary,
        column_config={
            "Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100
            ),
            "Số Lượng": st.column_config.NumberColumn("Số Lượng", format="%d"),
        },
        use_container_width=True,
        hide_index=True,
    )

  t_col3, t_col4 = st.columns([1.0, 1.0])
  with t_col3:
    st.markdown(f"##### 3. Chi Tiết Sản Lượng Dòng SP ({sel_fam})")
    st.dataframe(
        df_family_summary,
        column_config={
            "Tỷ Lệ Đóng Góp Tháng (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Đóng Góp Tháng",
                format="%.1f%%",
                min_value=0,
                max_value=100,
            ),
            f"SL Sản Xuất ({sel_fam})": st.column_config.NumberColumn(
                f"SL Sản Xuất ({sel_fam})", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )

  with t_col4:
    st.markdown("##### 4. Tổng Sản Lượng Cả Năm Các Mã Đầu 5")
    st.dataframe(
        df_year_summary,
        column_config={
            "Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100
            ),
            "Tổng SL Hoàn Thành Cả Năm": st.column_config.NumberColumn(
                "Tổng SL Cả Năm", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )


# ================= 8. RENDER NỘI DUNG CÁC TAB =================
with tab_co_khi:
  render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
with tab_tuti:
  render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")
with tab_cong_to:
  render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")