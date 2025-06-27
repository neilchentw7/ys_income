
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib
import platform

# 自動設定中文字型
if platform.system() == 'Windows':
    matplotlib.rcParams['font.family'] = 'Microsoft JhengHei'
elif platform.system() == 'Darwin':
    matplotlib.rcParams['font.family'] = 'Heiti TC'
else:
    matplotlib.rcParams['font.family'] = 'Noto Sans CJK TC'
matplotlib.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title="應收帳款分析", layout="wide")
st.title("📊 應收帳款視覺化分析")

uploaded_file = st.file_uploader("請上傳包含『應收帳款表』、『銷售日報輸入』與『銷售月報』分頁的 Excel 檔", type=["xlsx"])
if uploaded_file:
    shipment_value = 0
    try:
        df_sales = pd.read_excel(uploaded_file, sheet_name="銷售月報", header=None)
        shipment_value = df_sales.iloc[9, 32]  # AG10
        st.metric("📦 本月出貨數量", f"{shipment_value:,.0f}")
    except Exception as e:
        st.warning(f"無法讀取銷售月報 AG10 出貨數量：{e}")

    df_raw = pd.read_excel(uploaded_file, sheet_name="應收帳款表", skiprows=2)
    df_raw.columns = ['序號', '客戶名稱', '金額', '營業稅', '本月應收款', '備註']
    df = df_raw.drop(columns=['序號', '備註']).dropna(subset=['客戶名稱'])
    df = df[df['客戶名稱'] != 0]
    for col in ['金額', '營業稅', '本月應收款']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna(subset=['金額'])

    total_receivable = df['本月應收款'].sum()
    avg_price = total_receivable / shipment_value if shipment_value > 0 else 0

    st.metric("💰 本月應收款總金額", f"{total_receivable:,.0f} 元")
    st.metric("🧮 稅後平均單價", f"{avg_price:,.2f} 元/單位")

    df_dispatch = pd.read_excel(uploaded_file, sheet_name="銷售日報輸入", header=None)
    last_col_idx = df_dispatch.shape[1] - 1
    dispatch_records = []
    current_customer = None
    for i in range(len(df_dispatch)):
        name = df_dispatch.iloc[i, 1]
        if pd.notna(name):
            current_customer = name.strip()
        value = df_dispatch.iloc[i, last_col_idx]
        if current_customer and pd.notna(value):
            dispatch_records.append((current_customer, value))
    df_dispatch_summary = pd.DataFrame(dispatch_records, columns=['客戶名稱', '出貨量'])
    df_dispatch_summary = df_dispatch_summary.groupby('客戶名稱', as_index=False)['出貨量'].sum()

    df_receivable = df_raw[['客戶名稱', '本月應收款']].copy()
    df_receivable = df_receivable.dropna(subset=['客戶名稱'])
    df_receivable = df_receivable[df_receivable['客戶名稱'] != 0]
    df_receivable['客戶名稱'] = df_receivable['客戶名稱'].str.strip()
    df_dispatch_summary['客戶名稱'] = df_dispatch_summary['客戶名稱'].str.strip()
    df_merged = pd.merge(df_receivable, df_dispatch_summary, on='客戶名稱', how='left')
    df_merged['出貨量'] = pd.to_numeric(df_merged['出貨量'], errors='coerce').fillna(0)
    df_merged['稅後平均單價'] = df_merged.apply(
        lambda row: row['本月應收款'] / row['出貨量'] if row['出貨量'] > 0 else None,
        axis=1
    )
    df_merged = df_merged.sort_values('本月應收款', ascending=False).reset_index(drop=True)

    # 將數值格式化為字串，讓 dataframe 可直接顯示
    df_display = df_merged.copy()
    df_display['本月應收款'] = df_display['本月應收款'].map(lambda x: f"{x:,.0f}")
    df_display['出貨量'] = df_display['出貨量'].map(lambda x: f"{x:,.1f}")
    df_display['稅後平均單價'] = df_display['稅後平均單價'].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")

    st.subheader("📋 所有客戶平均單價分析表")
    st.dataframe(df_display, use_container_width=True)

    df_all = df.copy()
    df_top5 = df_all.sort_values("本月應收款", ascending=False).head(5)
    df_above = df_all[df_all["本月應收款"] >= 200000]
    df_below = df_all[df_all["本月應收款"] < 200000]

    def plot_bar(data, title, xtick_scale=0.7, label_fontscale=1.0):
        fig, ax = plt.subplots(figsize=(12, 6))
        data_sorted = data.sort_values("本月應收款", ascending=False).reset_index(drop=True)
        sns.barplot(data=data_sorted, x="客戶名稱", y="本月應收款", ax=ax)
        ax.set_title(title)
        for label in ax.get_xticklabels():
            label.set_rotation(45)
            label.set_horizontalalignment("right")
            label.set_fontsize(ax.xaxis.get_ticklabels()[0].get_size() * xtick_scale)
        for idx, row in data_sorted.iterrows():
            ax.text(idx, row["本月應收款"], f"{row['本月應收款']:,.0f}",
                    ha='center', va='bottom', fontsize=9 * label_fontscale)
        st.pyplot(fig)

    st.subheader("🔹 前五大客戶 - 本月應收帳款")
    plot_bar(df_top5, "前五大客戶 - 本月應收帳款")

    st.subheader("🔹 應收帳款 >= 20 萬 - 重點客戶")
    plot_bar(df_above, "應收帳款 >= 20 萬 - 客戶")

    st.subheader("🔹 應收帳款 < 20 萬 - 小額客戶")
    plot_bar(df_below, "應收帳款 < 20 萬 - 客戶", label_fontscale=0.5)
