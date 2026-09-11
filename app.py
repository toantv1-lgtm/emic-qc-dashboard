from datetime import date, datetime
import os
import sqlite3
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="EMIC QC Dashboard", page_icon="📊", layout="wide")

plt.rcParams["font.family"] = "sans-serif"
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
  SURFACE = "#FFFFFF"
  PRIMARY = "#3B82F6"
  SUCCESS = "#10B981"
  WARNING = "#F59E0B"
  DANGER = "#EF4444"
  PURPLE = "#A855F7"
  TEXT_PRIMARY = "#0F172A"


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


# --- THANH LỌC BÊN TRÁI ---
st.sidebar.title("🔍 Lọc Dữ Liệu Báo Cáo")
col_tu, col_den = st.sidebar.columns(2)
with col_tu:
  tu_date = st.date_input("Từ ngày", date(datetime.now().year, 1, 1))
with col_den:
  den_date = st.date_input("Đến ngày", date.today())

if st.sidebar.button("🔄 Làm mới dữ liệu"):
  st.cache_data.clear()

df_qa32, df_coois = load_data(tu_date, den_date)

st.title("📊 EMIC QC - HỆ THỐNG BÁO CÁO DỮ LIỆU TỰ ĐỘNG CHUYÊN SÂU")

tab_vt, tab_ck, tab_tuti, tab_ct = st.tabs([
    "📋 Báo Cáo Vật Tư",
    "⚙️ Báo Cáo Cơ Khí",
    "🔌 Báo Cáo TU/TI",
    "⚡ Báo Cáo Công Tơ",
])

# ================= TAB 1: VẬT TƯ =================
with tab_vt:
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

    col1, col2 = st.columns([2.5, 1])
    with col1:
      fig1, ax1 = plt.subplots(figsize=(10, 4.2), dpi=100)
      ax2 = ax1.twinx()
      x, w = np.arange(12), 0.38

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
      ax1.set_xticklabels(months_labels, fontweight="bold")
      ax1.set_ylabel(
          "← Số Lượng Lệnh (Trục Trái)", fontweight="bold", color=Theme.SUCCESS
      )
      ax2.set_ylabel(
          "Số Lượng Vật Tư / Mẫu [Log] (Trục Phải) →",
          fontweight="bold",
          color=Theme.PRIMARY,
      )
      ax1.set_title(
          "BÁO CÁO SỐ LƯỢNG LỆNH KIỂM & TỔNG VẬT TƯ VỀ / SỐ MẪU KIỂM",
          fontweight="bold",
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
          ncol=3,
      )
      st.pyplot(fig1)

    with col2:
      fig_pie, ax_pie = plt.subplots(figsize=(4, 4.2), dpi=100)
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
            wedgeprops=dict(width=0.45, edgecolor="white"),
        )
        texts[0].set_color(Theme.SUCCESS)
        texts[0].set_fontweight("bold")
        if len(texts) > 1:
          texts[1].set_color(Theme.DANGER)
          texts[1].set_fontweight("bold")
        ax_pie.text(
            0,
            0,
            f"TỔNG VẬT TƯ VỀ\n{total_ft_all:,.0f}",
            ha="center",
            va="center",
            fontweight="bold",
        )
      else:
        ax_pie.text(0, 0, "Chưa có dữ liệu", ha="center")
      ax_pie.set_title(
          "TỶ LỆ VẬT TƯ ĐẠT VS BỊ BLOCK LỖI", fontweight="bold"
      )
      st.pyplot(fig_pie)

    st.subheader("🏆 DANH SÁCH VẬT TƯ BỊ BLOCK & UD02, UD03")
    sorted_blocks = sorted(
        top_block_dict.values(),
        key=lambda x: (x["ud03"] + x["ud02"], x["ca_block"], x["ft_total"]),
        reverse=True,
    )
    if sorted_blocks:
      df_block = pd.DataFrame(sorted_blocks)
      df_block["SL Block / SL Về"] = df_block.apply(
          lambda r: f"{r['ca_block']:,.0f} / {r['ft_total']:,.0f}", axis=1
      )
      df_block = df_block[[
          "ma_vt",
          "ten_vt",
          "ncc",
          "ud02",
          "ud03",
          "SL Block / SL Về",
      ]]
      df_block.columns = [
          "Mã Vật Tư",
          "Tên Vật Tư",
          "Nhà Cung Cấp",
          "Lượt UD 02",
          "Lượt UD 03",
          "Tổng SL Block (CA) / SL Về",
      ]
      st.dataframe(df_block, use_container_width=True, hide_index=True)
    else:
      st.success("Không có vật tư nào bị Block hoặc UD 02, 03")


# ================= HÀM CHUNG CHO CÁC TAB COOIS =================
def render_coois_section(phan_he_code, title_text):
  df_sub = (
      df_coois[df_coois["phan_he"] == phan_he_code]
      if not df_coois.empty
      else pd.DataFrame()
  )

  if df_sub.empty:
    st.info(f"💡 Chưa có dữ liệu lệnh sản xuất cho phân hệ {title_text}.")
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

  col1, col2 = st.columns([2.5, 1])
  with col1:
    fig1, ax1 = plt.subplots(figsize=(8, 3.5), dpi=100)
    ax2 = ax1.twinx()
    x, w = np.arange(12), 0.35

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
      if t_qty > 0:
        pct = (m_comp_qty[i] / t_qty) * 100
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
        f"TIẾN ĐỘ SẢN XUẤT - {title_text}", fontweight="bold", fontsize=10
    )
    ax1.set_ylabel("← Tổng Lệnh", fontweight="bold", color=Theme.PRIMARY)
    ax2.set_ylabel("Số Lượng Giao [Log] →", fontweight="bold", color=Theme.SUCCESS)
    ax1.set_xticks(x)
    ax1.set_xticklabels(months_labels, fontweight="bold")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        frameon=False,
        ncol=4,
    )
    st.pyplot(fig1)

  with col2:
    fig2, ax3 = plt.subplots(figsize=(4, 3.5), dpi=100)
    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    if tot_qty_all > 0:
      pct_deliv = (deliv_qty_all / tot_qty_all) * 100
      pct_rem = 100.0 - pct_deliv
      wedges, texts = ax3.pie(
          [deliv_qty_all, rem_qty_all],
          labels=[
              f"Hoàn thành\n{pct_deliv:.1f}%\n({int(deliv_qty_all):,})",
              f"Chưa xong\n{pct_rem:.1f}%\n({int(rem_qty_all):,})",
          ],
          colors=[Theme.SUCCESS, Theme.WARNING],
          startangle=140,
          wedgeprops=dict(width=0.42, edgecolor="white"),
      )
      texts[0].set_color(Theme.SUCCESS)
      texts[0].set_fontweight("bold")
      if len(texts) > 1:
        texts[1].set_color(Theme.DANGER)
        texts[1].set_fontweight("bold")
      ax3.text(
          0,
          0,
          f"TỔNG KẾ HOẠCH\n{int(tot_qty_all):,}",
          ha="center",
          va="center",
          fontweight="bold",
      )
    else:
      ax3.text(0, 0, "Chưa có dữ liệu", ha="center")
    ax3.set_title("TỶ LỆ HOÀN THÀNH TỔNG QUAN", fontweight="bold", fontsize=10)
    st.pyplot(fig2)

  sub_5 = df_sub[
      df_sub["ma_tp"]
      .astype(str)
      .str.split(".")
      .str[0]
      .str.lstrip("0")
      .str.startswith("5")
  ].copy()
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
  available_fams = ["Tất cả dòng sản phẩm"] + sorted(list(set(raw_fams)))

  col3, col4 = st.columns(2)
  with col3:
    sel_fam = st.selectbox(
        "🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}"
    )

    m3_qty = [0.0] * 12
    if not sub_5.empty:
      sub_5["month"] = pd.to_datetime(
          sub_5["ngay_lenh_dt"], errors="coerce"
      ).dt.month
      sub_filtered = (
          sub_5[sub_5["mat_prefix"] == sel_fam]
          if sel_fam != "Tất cả dòng sản phẩm"
          else sub_5
      )
      for _, r in sub_filtered.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          m3_qty[int(m_val) - 1] += float(r["sl_ht"])

    fig3, ax_b3 = plt.subplots(figsize=(6, 3), dpi=100)
    ax_b3.bar(x, m3_qty, width=0.42, color=Theme.PRIMARY)
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

    ax_b3.set_xticks(x)
    ax_b3.set_xticklabels(months_labels, fontweight="bold")
    ax_b3.set_ylabel("SL Hoàn Thành [Log]", fontweight="bold", color=Theme.PRIMARY)
    ax_b3.set_title(
        f"SẢN LƯỢNG - DÒNG: {sel_fam}", fontweight="bold", fontsize=10
    )
    st.pyplot(fig3)

  with col4:
    fig4, ax_b4 = plt.subplots(figsize=(6, 3), dpi=100)
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
      if fams_x:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x = summary_fams["mat_prefix"].tolist()
        deliv_fams = summary_fams["sl_ht"].values
      else:
        fams_x, deliv_fams = ["Trống"], [0.0]
    else:
      fams_x, deliv_fams = ["Không có SP"], [0.0]

    x_b4 = np.arange(len(fams_x))
    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    ax_b4.bar(x_b4, deliv_fams, width=0.45, color=bar_colors)
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

    ax_b4.set_xticks(x_b4)
    ax_b4.set_xticklabels(fams_x, fontweight="bold")
    ax_b4.set_ylabel("Số Lượng SP [Log]", fontweight="bold", color=Theme.SUCCESS)
    ax_b4.set_title(
        "TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5", fontweight="bold", fontsize=10
    )
    st.pyplot(fig4)


with tab_ck:
  render_coois_section("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
with tab_tuti:
  render_coois_section("TU_TI", "🔌 BÁO CÁO TU/TI (LỆNH 3011)")
with tab_ct:
  render_coois_section("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")