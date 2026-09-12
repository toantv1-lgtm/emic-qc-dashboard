import os
import re
import sqlite3
from datetime import date, datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
import streamlit as st

# ================= 1. CẤU HÌNH TRÀN MÀN HÌNH 100% & THEME CSS =================
st.set_page_config(
    page_title="EMIC QC - HỆ THỐNG BÁO CÁO DỮ LIỆU TỰ ĐỘNG CHUYÊN SÂU",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .stApp {
            background-color: #F1F5F9;
            font-family: system-ui, -apple-system, sans-serif;
        }
        .main .block-container {
            max-width: 100% !important;
            padding: 1rem 1.5rem !important;
        }
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 1px solid #CBD5E1;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            background-color: #E2E8F0;
            padding: 4px;
            border-radius: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent;
            border-radius: 6px;
            padding: 6px 16px;
            border: none;
            font-weight: 600;
            color: #475569;
            font-size: 13px;
        }
        .stTabs [aria-selected="true"] {
            background-color: #3B82F6 !important;
            color: #FFFFFF !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        div[data-testid="stColumn"] {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            padding: 10px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .header-title {
            font-size: 20px;
            font-weight: 700;
            color: #0F172A;
            margin-bottom: 2px;
        }
        .sub-title {
            font-size: 12px;
            color: #475569;
            margin-bottom: 10px;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Cấu hình Font tương thích tốt trên môi trường Linux Server (Khắc phục log findfont)
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [
    "DejaVu Sans",
    "Liberation Sans",
    "Arial",
    "sans-serif",
]
plt.rcParams["font.size"] = 8.5
plt.rcParams["axes.unicode_minus"] = False

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
    "#6366F1",
    "#D97706",
    "#059669",
    "#7C3AED",
    "#E11D48",
]


class Theme:
  APP_BG = "#F1F5F9"
  SURFACE = "#FFFFFF"
  BORDER = "#CBD5E1"
  PRIMARY = "#3B82F6"
  SUCCESS = "#10B981"
  WARNING = "#F59E0B"
  DANGER = "#EF4444"
  PURPLE = "#A855F7"
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


# ================= 2. THANH BỘ LỌC DỮ LIỆU (SIDEBAR) =================
st.sidebar.title("🔍 Lọc Báo Cáo")
col_tu, col_den = st.sidebar.columns(2)
with col_tu:
  tu_date = st.date_input("Từ ngày", date(datetime.now().year, 1, 1))
with col_den:
  den_date = st.date_input("Đến ngày", date.today())

st.sidebar.markdown("---")

# Thay use_container_width=True bằng width="stretch"
if st.sidebar.button("🔄 Cập Nhật Lại Báo Cáo", width="stretch"):
  st.cache_data.clear()
  st.rerun()

df_qa32, df_coois = load_data(tu_date, den_date)

# ================= 3. TIÊU ĐỀ TRANG =================
st.markdown(
    '<div class="header-title">EMIC QC - HỆ THỐNG BÁO CÁO DỮ LIỆU TỰ ĐỘNG CHUYÊN'
    " SÂU</div>",
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Báo cáo dữ liệu tổng hợp kho vật tư và tiến độ lệnh'
    f" sản xuất từ {tu_date.strftime('%d.%m.%Y')} đến"
    f" {den_date.strftime('%d.%m.%Y')}</div>",
    unsafe_allow_html=True,
)

tab_vat_tu, tab_co_khi, tab_tuti, tab_cong_to = st.tabs([
    "📋 Báo Cáo Vật Tư",
    "⚙️ Báo Cáo Cơ Khí",
    "🔌 Báo Cáo TU/TI",
    "⚡ Báo Cáo Công Tơ",
])

# ================= 4. TAB BÁO CÁO VẬT TƯ =================
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
        if "VIHA" in ncc_str.upper():
          key = (
              "Mặt số công tơ",
              ncc_str if ncc_str else "Cty TNHH CN VIHA",
          )
          ma_display, ten_display = "Mặt số công tơ", "Mặt số công tơ"
        else:
          key = (ma_vt_str, ncc_str)
          ma_display, ten_display = ma_vt_str, ten_vt_str

        if key not in top_block_dict:
          top_block_dict[key] = {
              "ma_vt": ma_display,
              "ten_vt": ten_display,
              "ncc": key[1],
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

    col1, col2 = st.columns([1.8, 1])

    with col1:
      fig1 = plt.figure(figsize=(10, 3.8), dpi=150)
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

      b_ud01 = ax1.bar(
          x, ud01_m, width=w, color=Theme.SUCCESS, label="UD 01 (Đạt)"
      )
      b_ud02 = ax1.bar(
          x,
          ud02_m,
          width=w,
          bottom=ud01_m,
          color=Theme.WARNING,
          label="UD 02 (Đặc nhượng)",
      )
      bot_03 = [ud01_m[i] + ud02_m[i] for i in range(12)]
      b_ud03 = ax1.bar(
          x,
          ud03_m,
          width=w,
          bottom=bot_03,
          color=Theme.DANGER,
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
              fontsize=8.5,
              color=Theme.PURPLE,
          )

      total_by_sample_m = [
          by_inspected_m[i] + by_uninspected_m[i] for i in range(12)
      ]
      line_ft = ax2.plot(
          x,
          ft_qty_m,
          color=Theme.PRIMARY,
          marker="o",
          linewidth=2,
          label="Tổng số hàng về (FT)",
      )
      line_by = ax2.plot(
          x,
          total_by_sample_m,
          color=Theme.PURPLE,
          marker="s",
          linewidth=2,
          linestyle="--",
          label="Số mẫu phải kiểm (BY)",
      )

      ax2.set_yscale("symlog", linthresh=100)
      ax2.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

      ax1.set_xticks(x)
      ax1.set_xticklabels(months_labels, fontweight="bold", fontsize=8.5)
      ax1.set_xlim(-0.6, 11.6)
      ax1.set_ylabel(
          "← Số Lượng Lệnh (Trục Trái)",
          fontweight="bold",
          color=Theme.SUCCESS,
          fontsize=9,
      )
      ax2.set_ylabel(
          "Số Lượng Vật Tư / Mẫu [Log] (Trục Phải) →",
          fontweight="bold",
          color=Theme.PRIMARY,
          fontsize=9,
      )
      ax1.set_title(
          "BÁO CÁO SỐ LƯỢNG LỆNH KIỂM & TỔNG VẬT TƯ VỀ / SỐ MẪU KIỂM",
          fontweight="bold",
          fontsize=10.5,
          color=Theme.TEXT_PRIMARY,
      )

      all_handles = [b_ud01, b_ud02, b_ud03, line_ft[0], line_by[0]]
      all_labels = [
          "UD 01 (Đạt)",
          "UD 02 (Đặc nhượng)",
          "UD 03 (Trả lại)",
          "Tổng số hàng về (FT)",
          "Số mẫu phải kiểm (BY)",
      ]
      ax1.legend(
          all_handles,
          all_labels,
          loc="upper center",
          bbox_to_anchor=(0.5, -0.18),
          frameon=False,
          fontsize=8.5,
          ncol=3,
      )

      fig1.subplots_adjust(
          top=0.88, bottom=0.25, left=0.08, right=0.92, wspace=0.18
      )
      st.pyplot(fig1, width="stretch")

    with col2:
      fig_pie = plt.figure(figsize=(5.5, 3.8), dpi=150)
      fig_pie.patch.set_facecolor(Theme.SURFACE)
      ax_pie = fig_pie.add_subplot(111)
      ax_pie.set_facecolor(Theme.SURFACE)

      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      if total_ft_all > 0:
        wedges, texts, autotexts = ax_pie.pie(
            [ok_cnt, total_ca_block],
            labels=[
                f"Vật tư Đạt\n{ok_cnt:,.0f}",
                f"Bị Block (Lỗi)\n{total_ca_block:,.0f}",
            ],
            colors=[Theme.SUCCESS, Theme.DANGER],
            autopct="%1.1f%%",
            startangle=140,
            pctdistance=0.68,
            labeldistance=1.12,
            radius=1.3,
            center=(0, -0.10),
            wedgeprops=dict(width=0.48, edgecolor="white", linewidth=2),
        )
        texts[0].set_color(Theme.SUCCESS)
        texts[0].set_fontweight("bold")
        texts[0].set_fontsize(8.5)
        if len(texts) > 1:
          texts[1].set_color(Theme.DANGER)
          texts[1].set_fontweight("bold")
          texts[1].set_fontsize(8.5)
        for at in autotexts:
          at.set_fontweight("bold")
          at.set_fontsize(8.5)

        ax_pie.text(
            0,
            -0.10,
            f"TỔNG VẬT TƯ VỀ\n{total_ft_all:,.0f}",
            ha="center",
            va="center",
            fontweight="bold",
            fontsize=9,
            color=Theme.TEXT_PRIMARY,
        )
      else:
        ax_pie.text(0, -0.10, "Chưa có dữ liệu", ha="center", fontsize=9.5)
        ax_pie.axis("off")

      ax_pie.set_title(
          "TỶ LỆ VẬT TƯ ĐẠT VS BỊ BLOCK LỖI",
          fontweight="bold",
          fontsize=10.5,
          color=Theme.TEXT_PRIMARY,
      )
      fig_pie.subplots_adjust(top=0.88, bottom=0.15, left=0.05, right=0.95)
      st.pyplot(fig_pie, width="stretch")

    st.markdown("---")
    st.markdown(
        "##### 🏆 DANH SÁCH VẬT TƯ BỊ BLOCK & UD02, UD03", unsafe_allow_html=True
    )
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
      st.dataframe(df_block, width="stretch", hide_index=True)
    else:
      st.success("🎉 Không có vật tư nào bị Block hoặc UD 02, 03")


# ================= 5. HÀM CHUNG CHO CÁC TAB COOIS =================
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

  col1, col2 = st.columns([1, 1])

  with col1:
    fig1, ax1 = plt.subplots(figsize=(6.5, 2.7), dpi=150)
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
        color=Theme.PRIMARY,
        label="Lệnh Hoàn Thành",
    )
    ax1.bar(
        x - w / 2,
        m_uncomp_orders,
        width=w,
        bottom=m_comp_orders,
        color=Theme.DANGER,
        label="Lệnh Chưa Xong",
    )

    ax2.bar(
        x + w / 2,
        m_comp_qty,
        width=w,
        color=Theme.SUCCESS,
        label="SL Hoàn Thành",
    )
    ax2.bar(
        x + w / 2,
        m_uncomp_qty,
        width=w,
        bottom=m_comp_qty,
        color=Theme.WARNING,
        label="SL Chưa Xong",
    )

    ax2.set_yscale("symlog", linthresh=100)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    for i in range(12):
      t_qty = m_comp_qty[i] + m_uncomp_qty[i]
      pct = (m_comp_qty[i] / t_qty * 100) if t_qty > 0 else 0
      if t_qty > 0:
        ax2.text(
            x[i] + w / 2,
            t_qty * 1.05,
            f"{pct:.0f}%",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=7.5,
            color=Theme.SUCCESS,
        )

    ax1.set_title(
        f"TIẾN ĐỘ SẢN XUẤT - {title_clean}",
        fontweight="bold",
        fontsize=10,
        color=Theme.TEXT_PRIMARY,
    )
    ax1.set_ylabel(
        "← Tổng Lệnh (Trục Trái)",
        fontweight="bold",
        color=Theme.PRIMARY,
        fontsize=8.5,
    )
    ax2.set_ylabel(
        "Số Lượng Giao [Log] (Trục Phải) →",
        fontweight="bold",
        color=Theme.SUCCESS,
        fontsize=8.5,
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
        bbox_to_anchor=(0.5, -0.22),
        frameon=False,
        fontsize=7.5,
        ncol=4,
    )

    fig1.subplots_adjust(top=0.86, bottom=0.25, left=0.12, right=0.88)
    st.pyplot(fig1, width="stretch")

  with col2:
    fig2, ax3 = plt.subplots(figsize=(6.5, 2.7), dpi=150)
    fig2.patch.set_facecolor(Theme.SURFACE)
    ax3.set_facecolor(Theme.SURFACE)
    ax3.set_aspect("equal")

    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    if tot_qty_all > 0:
      wedges, texts = ax3.pie(
          [deliv_qty_all, rem_qty_all],
          labels=[
              (
                  f"Hoàn thành\n{pct_deliv:.1f}%\n({int(deliv_qty_all):,})"
              ),
              f"Chưa xong\n{pct_rem:.1f}%\n({int(rem_qty_all):,})",
          ],
          colors=[Theme.SUCCESS, Theme.WARNING],
          startangle=140,
          pctdistance=0.65,
          labeldistance=1.1,
          radius=1.1,
          wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
      )
      texts[0].set_color(Theme.SUCCESS)
      texts[0].set_fontweight("bold")
      texts[0].set_fontsize(8.5)
      if len(texts) > 1:
        texts[1].set_color(Theme.DANGER)
        texts[1].set_fontweight("bold")
        texts[1].set_fontsize(8.5)

      ax3.text(
          0,
          0,
          f"TỔNG KẾ HOẠCH\n{int(tot_qty_all):,}",
          ha="center",
          va="center",
          fontweight="bold",
          fontsize=9,
          color=Theme.TEXT_PRIMARY,
      )
    else:
      ax3.text(0, 0, "Chưa có dữ liệu", ha="center", fontsize=9)
      ax3.axis("off")

    ax3.set_title(
        "TỶ LỆ HOÀN THÀNH TỔNG QUAN",
        fontweight="bold",
        fontsize=10,
        color=Theme.TEXT_PRIMARY,
    )
    fig2.subplots_adjust(top=0.86, bottom=0.12, left=0.12, right=0.88)
    st.pyplot(fig2, width="stretch")

  col3, col4 = st.columns([1, 1])

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

  with col3:
    sel_fam = st.selectbox(
        "🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}"
    )

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

    fig3, ax_b3 = plt.subplots(figsize=(6.5, 2.7), dpi=150)
    fig3.patch.set_facecolor(Theme.SURFACE)
    ax_b3.set_facecolor(Theme.SURFACE)

    for spine in ["top"]:
      ax_b3.spines[spine].set_visible(False)
    ax_b3_right = ax_b3.twinx()
    for spine in ["top"]:
      ax_b3_right.spines[spine].set_visible(False)

    ax_b3.bar(
        x,
        m3_qty,
        width=0.42,
        color=Theme.PRIMARY,
        alpha=0.9,
        label="SL Sản Xuất",
    )
    ax_b3.set_yscale("symlog", linthresh=100)
    ax_b3.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    for i in range(12):
      v = m3_qty[i]
      if v > 0:
        ax_b3.text(
            x[i],
            v * 1.25 if v > 10 else v + 2,
            f"{int(v):,}",
            ha="center",
            va="bottom",
            fontweight="bold",
            fontsize=7,
            color=Theme.PRIMARY,
        )

    ax_b3_right.plot(
        x,
        defect_rate_m,
        color=Theme.DANGER,
        marker="o",
        linewidth=1.2,
        label="Tỷ lệ sai hỏng (%)",
    )

    ax_b3.set_xticks(x)
    ax_b3.set_xticklabels(months_labels, fontweight="bold", fontsize=8)
    ax_b3.set_xlim(-0.6, 11.6)
    ax_b3.set_ylabel(
        "SL Hoàn Thành [Log]", fontweight="bold", color=Theme.PRIMARY, fontsize=8.5
    )
    ax_b3_right.set_ylabel(
        "Sai Hỏng (%)", fontweight="bold", color=Theme.DANGER, fontsize=8.5
    )
    ax_b3_right.set_ylim(-1.0, 5.0)
    ax_b3.set_title(
        f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}",
        fontweight="bold",
        fontsize=10,
        color=Theme.TEXT_PRIMARY,
    )

    fig3.subplots_adjust(top=0.86, bottom=0.25, left=0.12, right=0.88)
    st.pyplot(fig3, width="stretch")

  with col4:
    fig4, ax_b4 = plt.subplots(figsize=(6.5, 2.7), dpi=150)
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

    for i in range(len(fams_x)):
      v = deliv_fams[i]
      if v > 0:
        ax_b4.text(
            x_b4[i],
            v * 1.25 if v > 10 else v + 2,
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
        color=Theme.DANGER,
        marker="s",
        linewidth=1.2,
        label="Tỷ lệ sai hỏng (%)",
    )

    ax_b4.set_xticks(x_b4)
    ax_b4.set_xticklabels(fams_x, fontweight="bold", fontsize=8.5)
    ax_b4.set_ylabel(
        "Số Lượng SP [Log]", fontweight="bold", color=Theme.SUCCESS, fontsize=8.5
    )
    ax_b4_right.set_ylabel(
        "Sai Hỏng (%)", fontweight="bold", color=Theme.DANGER, fontsize=8.5
    )
    ax_b4_right.set_ylim(-1.0, 5.0)
    ax_b4.set_title(
        "TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5",
        fontweight="bold",
        fontsize=10,
        color=Theme.TEXT_PRIMARY,
    )

    fig4.subplots_adjust(top=0.86, bottom=0.25, left=0.12, right=0.88)
    st.pyplot(fig4, width="stretch")


# ================= 6. RENDER NỘI DUNG CÁC TAB =================
with tab_co_khi:
  render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
with tab_tuti:
  render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")
with tab_cong_to:
  render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")