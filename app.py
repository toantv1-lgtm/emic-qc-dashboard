import re
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd
import streamlit as st

# ================= CẤU HÌNH FONT CHỮ HIỆN ĐẠI & ĐỒNG BỘ =================
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = [
    'Segoe UI',
    'Inter',
    'Roboto',
    'Arial',
    'DejaVu Sans',
]
plt.rcParams['axes.edgecolor'] = '#CBD5E1'
plt.rcParams['axes.linewidth'] = 0.8


def clean_emoji(text):
  """Lọc Emoji để tránh lỗi vỡ font trên Matplotlib."""
  return re.sub(r'[^\w\s\(\)\-\/\.\,\:]', '', str(text)).strip()


# ================= HÀM HIỂN THỊ BÁO CÁO CƠ KHÍ / TU-TI / CÔNG TƠ =================
def render_coois_tab_layout(phan_he_code, title_text):
  df_sub = (
      df_coois[df_coois['phan_he'] == phan_he_code]
      if not df_coois.empty
      else pd.DataFrame()
  )

  if df_sub.empty:
    st.info(f'💡 Chưa có dữ liệu lệnh sản xuất cho phân hệ {title_text}.')
    return

  months_labels = [f'T{i}' for i in range(1, 13)]
  m_comp_qty, m_uncomp_qty = [0.0] * 12, [0.0] * 12
  m_tot_orders, m_uncomp_orders = [0] * 12, [0] * 12
  tot_qty_all, deliv_qty_all = 0.0, 0.0

  for _, r in df_sub.iterrows():
    try:
      m_idx = (
          datetime.strptime(str(r['ngay_lenh_dt']).split()[0], '%Y-%m-%d').month
          - 1
      )
    except:
      m_idx = 0
    if not (0 <= m_idx < 12):
      m_idx = 0

    sl_t, sl_h = float(r['sl_tong']), float(r['sl_ht'])
    uncomp_q = max(0.0, sl_t - sl_h)

    tot_qty_all += sl_t
    deliv_qty_all += sl_h
    m_comp_qty[m_idx] += sl_h
    m_uncomp_qty[m_idx] += uncomp_q

    m_tot_orders[m_idx] += 1
    if sl_h < sl_t:
      m_uncomp_orders[m_idx] += 1

  clean_title = clean_emoji(title_text)

  # ================= HÀNG 1: TIẾN ĐỘ SẢN XUẤT & BÁO CÁO TỔNG QUAN =================
  col1, col2 = st.columns([1.6, 1])

  # --- 1. BIỂU ĐỒ TIẾN ĐỘ SẢN XUẤT (TRÁI) ---
  with col1:
    fig1, ax1 = plt.subplots(figsize=(7.0, 3.2), dpi=180)
    fig1.patch.set_facecolor(Theme.SURFACE)
    ax1.set_facecolor(Theme.SURFACE)
    ax2 = ax1.twinx()

    for spine in ['top']:
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
        label='Lệnh Hoàn Thành',
    )
    ax1.bar(
        x - w / 2,
        m_uncomp_orders,
        width=w,
        bottom=m_comp_orders,
        color=Theme.DANGER,
        label='Lệnh Chưa Xong',
    )

    ax2.bar(
        x + w / 2,
        m_comp_qty,
        width=w,
        color=Theme.SUCCESS,
        label='SL Hoàn Thành',
    )
    ax2.bar(
        x + w / 2,
        m_uncomp_qty,
        width=w,
        bottom=m_comp_qty,
        color=Theme.WARNING,
        label='SL Chưa Xong',
    )

    ax2.set_yscale('symlog', linthresh=100)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    # Tăng khoảng không trần (ylim) để con số % không chạm viền trên
    max_q = max([m_comp_qty[i] + m_uncomp_qty[i] for i in range(12)] + [100])
    ax2.set_ylim(0, max_q * 3.2)

    for i in range(12):
      t_qty = m_comp_qty[i] + m_uncomp_qty[i]
      pct = (m_comp_qty[i] / t_qty * 100) if t_qty > 0 else 0
      if t_qty > 0:
        ax2.text(
            x[i] + w / 2,
            t_qty * 1.12,
            f'{pct:.0f}%',
            ha='center',
            va='bottom',
            fontweight='bold',
            fontsize=7.5,
            color=Theme.SUCCESS,
        )

    ax1.set_title(
        f'TIẾN ĐỘ SẢN XUẤT - {clean_title}',
        fontweight='bold',
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )
    ax1.set_ylabel(
        '← Tổng Lệnh', fontweight='bold', color=Theme.PRIMARY, fontsize=8
    )
    ax2.set_ylabel(
        'Số Lượng Giao [Log] →',
        fontweight='bold',
        color=Theme.SUCCESS,
        fontsize=8,
    )
    ax1.set_xticks(x)
    ax1.set_xticklabels(months_labels, fontweight='bold', fontsize=8)
    ax1.tick_params(axis='both', labelsize=8)
    ax2.tick_params(axis='y', labelsize=8)
    ax1.set_xlim(-0.6, 11.6)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(
        lines1 + lines2,
        labels1 + labels2,
        loc='upper center',
        bbox_to_anchor=(0.5, -0.18),
        frameon=False,
        fontsize=7.5,
        ncol=4,
    )

    fig1.subplots_adjust(top=0.88, bottom=0.22, left=0.10, right=0.90)
    st.pyplot(fig1, use_container_width=True)

  # --- 2. BIỂU ĐỒ TRÒN DONUT CHART (PHẢI) ---
  with col2:
    fig2, ax3 = plt.subplots(figsize=(4.2, 3.2), dpi=180)
    fig2.patch.set_facecolor(Theme.SURFACE)
    ax3.set_facecolor(Theme.SURFACE)

    rem_qty_all = max(0.0, tot_qty_all - deliv_qty_all)
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    if tot_qty_all > 0:
      wedges, texts = ax3.pie(
          [deliv_qty_all, rem_qty_all],
          labels=[
              f'Hoàn thành\n{pct_deliv:.1f}%\n({int(deliv_qty_all):,})',
              f'Chưa xong\n{pct_rem:.1f}%\n({int(rem_qty_all):,})',
          ],
          colors=[Theme.SUCCESS, Theme.WARNING],
          startangle=140,
          pctdistance=0.6,
          labeldistance=1.18,
          radius=0.78,  # Thu nhỏ bán kính để chữ không bị tràn viền
          wedgeprops=dict(width=0.32, edgecolor='white', linewidth=2),
      )

      # Đảm bảo phông chữ vừa vặn, chuẩn kích thước
      texts[0].set_color(Theme.SUCCESS)
      texts[0].set_fontweight('bold')
      texts[0].set_fontsize(7.5)
      if len(texts) > 1:
        texts[1].set_color(Theme.DANGER)
        texts[1].set_fontweight('bold')
        texts[1].set_fontsize(7.5)

      ax3.text(
          0,
          0,
          f'TỔNG KẾ HOẠCH\n{int(tot_qty_all):,}',
          ha='center',
          va='center',
          fontweight='bold',
          fontsize=8.5,
          color=Theme.TEXT_PRIMARY,
      )
    else:
      ax3.text(
          0,
          0,
          'Chưa có dữ liệu',
          ha='center',
          fontsize=8.5,
          color=Theme.TEXT_PRIMARY,
      )
      ax3.axis('off')

    ax3.set_title(
        'TỶ LỆ HOÀN THÀNH TỔNG QUAN',
        fontweight='bold',
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )
    fig2.subplots_adjust(top=0.88, bottom=0.10, left=0.10, right=0.90)
    st.pyplot(fig2, use_container_width=True)

  # ================= HÀNG 2: BỘ LỌC DÒNG SP VÀ 2 BIỂU ĐỒ SẢN LƯỢNG =================
  st.markdown('---')

  sub_5 = (
      df_sub[
          df_sub['ma_tp']
          .astype(str)
          .str.split('.')
          .str[0]
          .str.lstrip('0')
          .str.startswith('5')
      ].copy()
      if not df_sub.empty
      else pd.DataFrame()
  )

  raw_fams = (
      [
          str(x).strip()
          for x in sub_5['mat_prefix'].unique()
          if pd.notna(x)
          and str(x).strip()
          and str(x).strip().lower() not in ['none', 'nan']
      ]
      if not sub_5.empty
      else []
  )
  available_fams = ['Tất cả dòng sản phẩm'] + sorted(list(set(raw_fams)))

  # ĐƯA BỘ LỌC RA TRƯỚC BẮT ĐẦU CỘT Để 2 BIỂU ĐỒ BÊN DƯỚI PHẲNG HÀNG 100%
  col_sel, _ = st.columns([1, 2])
  with col_sel:
    sel_fam = st.selectbox(
        '🎯 Chọn Dòng SP (Đầu 5):', available_fams, key=f'cb_{phan_he_code}'
    )

  col3, col4 = st.columns([1, 1])

  # --- 3. BIỂU ĐỒ SẢN LƯỢNG THEO DÒNG SP (DƯỚI TRÁI) ---
  with col3:
    m3_qty = [0.0] * 12
    if not sub_5.empty:
      sub_5_df = sub_5.copy()
      sub_5_df['month'] = pd.to_datetime(
          sub_5_df['ngay_lenh_dt'], errors='coerce'
      ).dt.month
      sub_filtered = (
          sub_5_df[sub_5_df['mat_prefix'] == sel_fam]
          if sel_fam != 'Tất cả dòng sản phẩm'
          else sub_5_df
      )
      for _, r in sub_filtered.iterrows():
        m_val = r['month']
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          m3_qty[int(m_val) - 1] += float(r['sl_ht'])

    defect_rate_m = [0.0] * 12

    fig3, ax_b3 = plt.subplots(figsize=(6.0, 2.7), dpi=180)
    fig3.patch.set_facecolor(Theme.SURFACE)
    ax_b3.set_facecolor(Theme.SURFACE)

    for spine in ['top']:
      ax_b3.spines[spine].set_visible(False)
    ax_b3_right = ax_b3.twinx()
    for spine in ['top']:
      ax_b3_right.spines[spine].set_visible(False)

    ax_b3.bar(
        x,
        m3_qty,
        width=0.42,
        color=Theme.PRIMARY,
        alpha=0.9,
        label='SL Sản Xuất',
    )
    ax_b3.set_yscale('symlog', linthresh=100)
    ax_b3.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    max_v_b3 = max(m3_qty + [100])
    ax_b3.set_ylim(0, max_v_b3 * 3.2)  # Khoảng trống trần vừa đủ

    for i in range(12):
      v = m3_qty[i]
      if v > 0:
        ax_b3.text(
            x[i],
            v * 1.15,
            f'{int(v):,}',
            ha='center',
            va='bottom',
            fontweight='bold',
            fontsize=7,
            color=Theme.PRIMARY,
        )

    ax_b3_right.plot(
        x,
        defect_rate_m,
        color=Theme.DANGER,
        marker='o',
        linewidth=1.2,
        label='Tỷ lệ sai hỏng (%)',
    )

    ax_b3.set_xticks(x)
    ax_b3.set_xticklabels(months_labels, fontweight='bold', fontsize=8)
    ax_b3.tick_params(axis='both', labelsize=8)
    ax_b3_right.tick_params(axis='y', labelsize=8)
    ax_b3.set_xlim(-0.6, 11.6)
    ax_b3.set_ylabel(
        'SL Hoàn Thành [Log]', fontweight='bold', color=Theme.PRIMARY, fontsize=8
    )
    ax_b3_right.set_ylabel(
        'Sai Hỏng (%)', fontweight='bold', color=Theme.DANGER, fontsize=8
    )
    ax_b3_right.set_ylim(-1.0, 5.0)
    ax_b3.set_title(
        f'SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}',
        fontweight='bold',
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )

    fig3.subplots_adjust(top=0.86, bottom=0.22, left=0.12, right=0.88)
    st.pyplot(fig3, use_container_width=True)

  # --- 4. BIỂU ĐỒ TỔNG SẢN LƯỢNG MÃ ĐẦU 5 (DƯỚI PHẢI) ---
  with col4:
    fig4, ax_b4 = plt.subplots(figsize=(6.0, 2.7), dpi=180)
    fig4.patch.set_facecolor(Theme.SURFACE)
    ax_b4.set_facecolor(Theme.SURFACE)

    for spine in ['top']:
      ax_b4.spines[spine].set_visible(False)
    ax_b4_right = ax_b4.twinx()
    for spine in ['top']:
      ax_b4_right.spines[spine].set_visible(False)

    if not sub_5.empty:
      summary_fams = (
          sub_5.groupby('mat_prefix')[['sl_ht']].sum().reset_index()
      )
      fams_x = [
          str(val)
          for val in summary_fams['mat_prefix'].tolist()
          if pd.notna(val)
          and str(val).strip()
          and str(val).strip().lower() not in ['none', 'nan']
      ]
      if not fams_x:
        fams_x, deliv_fams, defect_fams = ['Trống'], [0.0], [0.0]
      else:
        summary_fams = summary_fams[summary_fams['mat_prefix'].isin(fams_x)]
        fams_x = summary_fams['mat_prefix'].tolist()
        deliv_fams = summary_fams['sl_ht'].values
        defect_fams = [0.0] * len(fams_x)
    else:
      fams_x, deliv_fams, defect_fams = ['Không có SP'], [0.0], [0.0]

    x_b4 = np.arange(len(fams_x))
    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    bars_deliv = ax_b4.bar(
        x_b4,
        deliv_fams,
        width=0.45,
        label='SL Hoàn Thành Cả Năm',
        color=bar_colors,
    )
    ax_b4.set_yscale('symlog', linthresh=100)
    ax_b4.yaxis.set_major_formatter(ticker.FuncFormatter(log_formatter))

    # Tăng headroom để số liệu lớn (như 117,651) không bị chạm viền trên
    max_v_b4 = max(list(deliv_fams) + [100])
    ax_b4.set_ylim(0, max_v_b4 * 3.5)

    for i in range(len(fams_x)):
      v = deliv_fams[i]
      if v > 0:
        ax_b4.text(
            x_b4[i],
            v * 1.18,
            f'{int(v):,}',
            ha='center',
            va='bottom',
            fontweight='bold',
            fontsize=7.5,
            color=bar_colors[i % len(bar_colors)],
        )

    ax_b4_right.plot(
        x_b4,
        defect_fams,
        color=Theme.DANGER,
        marker='s',
        linewidth=1.2,
        label='Tỷ lệ sai hỏng (%)',
    )

    ax_b4.set_xticks(x_b4)
    ax_b4.set_xticklabels(fams_x, fontweight='bold', fontsize=8)
    ax_b4.tick_params(axis='both', labelsize=8)
    ax_b4_right.tick_params(axis='y', labelsize=8)
    ax_b4.set_ylabel(
        'Số Lượng SP [Log]', fontweight='bold', color=Theme.SUCCESS, fontsize=8
    )
    ax_b4_right.set_ylabel(
        'Sai Hỏng (%)', fontweight='bold', color=Theme.DANGER, fontsize=8
    )
    ax_b4_right.set_ylim(-1.0, 5.0)
    ax_b4.set_title(
        'TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5',
        fontweight='bold',
        fontsize=9.5,
        color=Theme.TEXT_PRIMARY,
        pad=10,
    )

    fig4.subplots_adjust(top=0.86, bottom=0.22, left=0.12, right=0.88)
    st.pyplot(fig4, use_container_width=True)