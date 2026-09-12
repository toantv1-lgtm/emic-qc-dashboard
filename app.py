from datetime import date, datetime
import os
import re
import sqlite3
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ================= 1. CẤU HÌNH GIAO DIỆN & CSS THU GỌN LỀ =================
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

        /* Nút Navigation Tabs Căn Giữa Đỉnh Trang */
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
        
        /* Card Khung Bao Cho Biểu Đồ */
        .chart-card {
            background-color: #FFFFFF;
            border-radius: 10px;
            border: 1px solid #CBD5E1;
            padding: 6px 10px;
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

# Bảng Màu Chuẩn Gốc 100%
COLOR_SUCCESS = "#10B981"  # Xanh lá (UD01 / SL Hoàn thành)
COLOR_PRIMARY = "#3B82F6"  # Xanh dương (Lệnh HT / Hàng về FT)
COLOR_DANGER = "#EF4444"  # Đỏ (UD03 / Bị Block / Lệnh chưa xong)
COLOR_WARNING = "#F59E0B"  # Cam (UD02 / SL chưa xong)
COLOR_PURPLE = "#A855F7"  # Tím (Mẫu BY)
COLOR_TEXT = "#0F172A"

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


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


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


# ================= 2. SIDEBAR BỘ LỌC NGÀY =================
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

# ================= 3. NAVIGATION TABS CĂN GIỮA ĐỈNH TRANG =================
tab_vat_tu, tab_co_khi, tab_tuti, tab_cong_to = st.tabs([
    "📋 Báo Cáo Vật Tư",
    "⚙️ Báo Cáo Cơ Khí",
    "🔌 Báo Cáo TU/TI",
    "⚡ Báo Cáo Công Tơ",
])

# ================= 4. TAB 1: BÁO CÁO VẬT TƯ =================
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
      fig1 = make_subplots(specs=[[{"secondary_y": True}]])
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud01_m,
              name="UD 01 (Đạt)",
              marker_color=COLOR_SUCCESS,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud02_m,
              name="UD 02 (Đặc nhượng)",
              marker_color=COLOR_WARNING,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud03_m,
              name="UD 03 (Trả lại)",
              marker_color=COLOR_DANGER,
          ),
          secondary_y=False,
      )

      total_by = [by_inspected_m[i] + by_uninspected_m[i] for i in range(12)]
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=total_by,
              name="Số mẫu phải kiểm (BY)",
              line=dict(color=COLOR_PURPLE, width=2, dash="dash"),
          ),
          secondary_y=True,
      )
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=ft_qty_m,
              name="Tổng số hàng về (FT)",
              line=dict(color=COLOR_PRIMARY, width=2),
          ),
          secondary_y=True,
      )

      fig1.update_layout(
          title=dict(
              text="BÁO CÁO SỐ LƯỢNG LỆNH KIỂM & TỔNG VẬT TƯ VỀ / SỐ MẪU KIỂM",
              font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
              x=0.5,
              xanchor="center",
          ),
          barmode="stack",
          margin=dict(l=10, r=10, t=35, b=10),
          height=260,  # Ép chuẩn chiều cao
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          legend=dict(
              orientation="h",
              yanchor="top",
              y=-0.18,
              xanchor="center",
              x=0.5,
              font=dict(size=9),
          ),
      )
      fig1.update_xaxes(showgrid=False, tickfont=dict(size=10))
      fig1.update_yaxes(
          title_text="← Số Lượng Lệnh",
          title_font=dict(size=10, color=COLOR_SUCCESS),
          secondary_y=False,
          showgrid=True,
          gridcolor="#E2E8F0",
      )
      fig1.update_yaxes(
          title_text="Số Lượng Vật Tư / Mẫu [Log] →",
          title_font=dict(size=10, color=COLOR_PRIMARY),
          type="log",
          secondary_y=True,
          showgrid=False,
      )

      st.plotly_chart(
          fig1, use_container_width=True, config={"displayModeBar": False}
      )
      st.markdown("</div>", unsafe_allow_html=True)

    with col2:
      st.markdown('<div class="chart-card">', unsafe_allow_html=True)
      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      fig2 = go.Figure(
          data=[
              go.Pie(
                  labels=["Vật tư Đạt", "Bị Block (Lỗi)"],
                  values=[ok_cnt, total_ca_block],
                  hole=0.55,
                  marker_colors=[COLOR_SUCCESS, COLOR_DANGER],
                  textinfo="label+percent",
                  insidetextorientation="radial",
              )
          ]
      )
      fig2.update_layout(
          title=dict(
              text="TỶ LỆ VẬT TƯ ĐẠT VS BỊ BLOCK LỖI",
              font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
              x=0.5,
              xanchor="center",
          ),
          margin=dict(l=10, r=10, t=35, b=10),
          height=260,  # Ép chuẩn chiều cao
          paper_bgcolor="rgba(0,0,0,0)",
          showlegend=False,
          annotations=[
              dict(
                  text=f"TỔNG VẬT TƯ VỀ<br><b>{int(total_ft_all):,}</b>",
                  x=0.5,
                  y=0.5,
                  font_size=10,
                  showarrow=False,
              )
          ],
      )
      st.plotly_chart(
          fig2, use_container_width=True, config={"displayModeBar": False}
      )
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


# ================= 5. HÀM CHUNG BÁO CÁO COOIS (PLOTLY CÂN ĐỐI 100%) =================
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

  # HÀNG 1: TIẾN ĐỘ SẢN XUẤT (CỘT 2.1) & DONUT CHART (CỘT 1.0)
  col1, col2 = st.columns([2.1, 1.0])

  with col1:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    m_comp_orders = [m_tot_orders[i] - m_uncomp_orders[i] for i in range(12)]

    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_comp_orders,
            name="Lệnh Hoàn Thành",
            marker_color=COLOR_PRIMARY,
        ),
        secondary_y=False,
    )
    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_uncomp_orders,
            name="Lệnh Chưa Xong",
            marker_color=COLOR_DANGER,
        ),
        secondary_y=False,
    )
    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_comp_qty,
            name="SL Hoàn Thành",
            marker_color=COLOR_SUCCESS,
        ),
        secondary_y=True,
    )
    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_uncomp_qty,
            name="SL Chưa Xong",
            marker_color=COLOR_WARNING,
        ),
        secondary_y=True,
    )

    fig1.update_layout(
        title=dict(
            text=f"TIẾN ĐỘ SẢN XUẤT - {title_clean}",
            font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
            x=0.5,
            xanchor="center",
        ),
        barmode="stack",
        margin=dict(l=10, r=10, t=35, b=10),
        height=260,  # Ép chuẩn chiều cao
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            font=dict(size=9),
        ),
    )
    fig1.update_xaxes(showgrid=False, tickfont=dict(size=10))
    fig1.update_yaxes(
        title_text="← Tổng Lệnh",
        title_font=dict(size=10, color=COLOR_PRIMARY),
        secondary_y=False,
        showgrid=True,
        gridcolor="#E2E8F0",
    )
    fig1.update_yaxes(
        title_text="Số Lượng Giao [Log] →",
        title_font=dict(size=10, color=COLOR_SUCCESS),
        type="log",
        secondary_y=True,
        showgrid=False,
    )

    st.plotly_chart(
        fig1, use_container_width=True, config={"displayModeBar": False}
    )
    st.markdown("</div>", unsafe_allow_html=True)

  with col2:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    fig2 = go.Figure(
        data=[
            go.Pie(
                labels=["Hoàn thành", "Chưa xong"],
                values=[deliv_qty_all, rem_qty_all],
                hole=0.55,
                marker_colors=[COLOR_SUCCESS, COLOR_WARNING],
                textinfo="label+percent",
                insidetextorientation="radial",
            )
        ]
    )
    fig2.update_layout(
        title=dict(
            text="TỶ LỆ HOÀN THÀNH TỔNG QUAN",
            font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
            x=0.5,
            xanchor="center",
        ),
        margin=dict(l=10, r=10, t=35, b=10),
        height=260,  # Ép chuẩn chiều cao
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        annotations=[
            dict(
                text=f"TỔNG KẾ HOẠCH<br><b>{int(tot_qty_all):,}</b>",
                x=0.5,
                y=0.5,
                font_size=10,
                showarrow=False,
            )
        ],
    )
    st.plotly_chart(
        fig2, use_container_width=True, config={"displayModeBar": False}
    )
    st.markdown("</div>", unsafe_allow_html=True)

  # HÀNG 2: BỘ LỌC ĐẶT NGOÀI TRÁNH Ộ TRẮNG THỪA & 2 BIỂU ĐỒ BÊN DƯỚI CÂN BẰNG
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

  # Bộ Lọc Dòng SP Đặt Trực Tiếp
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

    fig3 = go.Figure()
    fig3.add_trace(
        go.Bar(
            x=months_labels,
            y=m3_qty,
            name="SL Sản Xuất",
            marker_color=COLOR_PRIMARY,
            text=[f"{int(v):,}" if v > 0 else "" for v in m3_qty],
            textposition="auto",
        )
    )
    fig3.update_layout(
        title=dict(
            text=f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}",
            font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
            x=0.5,
            xanchor="center",
        ),
        margin=dict(l=10, r=10, t=35, b=10),
        height=260,  # Ép chuẩn chiều cao
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
    )
    fig3.update_xaxes(showgrid=False, tickfont=dict(size=10))
    fig3.update_yaxes(
        title_text="SL Hoàn Thành [Log]",
        type="log",
        showgrid=True,
        gridcolor="#E2E8F0",
    )

    st.plotly_chart(
        fig3, use_container_width=True, config={"displayModeBar": False}
    )
    st.markdown("</div>", unsafe_allow_html=True)

  # --- DƯỚI PHẢI: TỔNG SẢN LƯỢNG MÃ ĐẦU 5 ---
  with col4:
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
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
        fams_x, deliv_fams = ["Trống"], [0.0]
      else:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x = summary_fams["mat_prefix"].tolist()
        deliv_fams = summary_fams["sl_ht"].values
    else:
      fams_x, deliv_fams = ["Không có SP"], [0.0]

    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    fig4 = go.Figure()
    fig4.add_trace(
        go.Bar(
            x=fams_x,
            y=deliv_fams,
            marker_color=bar_colors,
            text=[f"{int(v):,}" if v > 0 else "" for v in deliv_fams],
            textposition="auto",
        )
    )
    fig4.update_layout(
        title=dict(
            text="TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5",
            font=dict(size=12, color=COLOR_TEXT, family="sans-serif"),
            x=0.5,
            xanchor="center",
        ),
        margin=dict(l=10, r=10, t=35, b=10),
        height=260,  # Ép chuẩn chiều cao
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
    )
    fig4.update_xaxes(showgrid=False, tickfont=dict(size=10))
    fig4.update_yaxes(
        title_text="Số Lượng SP [Log]",
        type="log",
        showgrid=True,
        gridcolor="#E2E8F0",
    )

    st.plotly_chart(
        fig4, use_container_width=True, config={"displayModeBar": False}
    )
    st.markdown("</div>", unsafe_allow_html=True)


# ================= 6. RENDER NỘI DUNG CÁC TAB =================
with tab_co_khi:
  render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
with tab_tuti:
  render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")
with tab_cong_to:
  render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")