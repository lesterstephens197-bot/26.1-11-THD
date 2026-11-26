import io
import pandas as pd
import plotly.express as px
import streamlit as st

# 页面基础配置
st.set_page_config(
    page_title="PO 订单销量与销售额看板", page_icon="📊", layout="wide"
)

st.title("📊 订单销量与销售额数据看板")
st.markdown("上传订单明细表（支持 `.xlsx` 或 `.csv`），自动生成月度及 Vendor SKU 维度的分析看板与可下载报表。")


# --- 数据清洗函数 ---
@st.cache_data
def load_and_clean_data(uploaded_file):
    # 读取文件
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # 去除表头首尾空格
    df.columns = df.columns.str.strip()

    # 检查必要字段是否存在
    required_cols = [
        "Order Date",
        "Vendor SKU",
        "Quantity",
        "Total Cost",
        "Line Status",
    ]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"上传的文件缺少必要列: {', '.join(missing_cols)}")
        return None

    # 1. 清洗日期并提取年月 (YYYY-MM)
    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
    df = df.dropna(subset=["Order Date"])
    df["Order Month"] = df["Order Date"].dt.to_period("M").astype(str)

    # 2. 清洗数值字段（去除可能存在的 $ 符号和千分位逗号）
    for col in ["Quantity", "Total Cost", "Unit Cost"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(r"[$,¥€\s]", "", regex=True)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 3. 填充空的文本字段
    df["Vendor SKU"] = df["Vendor SKU"].fillna("Unknown SKU").astype(str)
    df["Line Status"] = df["Line Status"].fillna("Unknown").astype(str)
    if "Description" in df.columns:
        df["Description"] = df["Description"].fillna("").astype(str)

    return df


# --- 导出 Excel 辅助函数 ---
def convert_df_to_excel(monthly_df, sku_qty_pivot, sku_sales_pivot, sku_detail_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        monthly_df.to_excel(writer, index=False, sheet_name="1_各月总览")
        sku_qty_pivot.to_excel(writer, sheet_name="2_SKU各月销量矩阵")
        sku_sales_pivot.to_excel(writer, sheet_name="3_SKU各月销售额矩阵")
        sku_detail_df.to_excel(writer, index=False, sheet_name="4_SKU月度明细表")
    return output.getvalue()


# --- 侧边栏：文件上传与筛选 ---
with st.sidebar:
    st.header("1. 上传数据")
    uploaded_file = st.file_uploader("选择 Excel 或 CSV 文件", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    df = load_and_clean_data(uploaded_file)

    if df is not None and not df.empty:
        # 侧边栏过滤器
        with st.sidebar:
            st.header("2. 数据筛选")
            
            # 订单状态筛选（默认全选，方便剔除 Cancelled 等状态）
            all_statuses = sorted(df["Line Status"].unique().tolist())
            selected_statuses = st.multiselect(
                "筛选 Line Status (订单状态)",
                options=all_statuses,
                default=all_statuses,
            )

            # 月份筛选
            all_months = sorted(df["Order Month"].unique().tolist())
            selected_months = st.multiselect(
                "筛选 Order Month (月份)",
                options=all_months,
                default=all_months,
            )

            # 货币展示（如果存在多种货币提示用户）
            currencies = df["Unit Cost Currency"].dropna().unique().tolist() if "Unit Cost Currency" in df.columns else []
            if len(currencies) > 1:
                st.warning(f"⚠️ 检测到多种货币: {', '.join(currencies)}，建议在下方筛选单一货币后再汇总。")
                selected_currency = st.multiselect("筛选货币", currencies, default=currencies)
                df = df[df["Unit Cost Currency"].isin(selected_currency)]

        # 应用筛选条件
        filtered_df = df[
            (df["Line Status"].isin(selected_statuses))
            & (df["Order Month"].isin(selected_months))
        ]

        if filtered_df.empty:
            st.warning("当前筛选条件下没有数据，请调整侧边栏筛选器。")
            st.stop()

        # --- 核心指标卡 (KPI) ---
        total_qty = int(filtered_df["Quantity"].sum())
        total_sales = filtered_df["Total Cost"].sum()
        total_orders = filtered_df["PO Number"].nunique() if "PO Number" in filtered_df.columns else len(filtered_df)
        active_skus = filtered_df["Vendor SKU"].nunique()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总销量 (Quantity)", f"{total_qty:,}")
        col2.metric("总销售额 (Total Cost)", f"{total_sales:,.2f}")
        col3.metric("PO 订单数", f"{total_orders:,}")
        col4.metric("活跃 Vendor SKU 数", f"{active_skus:,}")

        st.divider()

        # ==========================================
        # 模块一：各月销量和销售额看板
        # ==========================================
        st.subheader("📅 1. 各月销量与销售额趋势")

        monthly_summary = (
            filtered_df.groupby("Order Month", as_index=False)
            .agg(
                总销量=("Quantity", "sum"),
                总销售额=("Total Cost", "sum"),
                订单行数=("Order Date", "count"),
            )
            .sort_values("Order Month")
        )

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            fig_qty = px.bar(
                monthly_summary,
                x="Order Month",
                y="总销量",
                text_auto=True,
                title="各月总销量 (Quantity)",
                color_discrete_sequence=["#2E86C1"],
            )
            fig_qty.update_layout(xaxis_title="月份", yaxis_title="销量")
            st.plotly_chart(fig_qty, use_container_width=True)

        with chart_col2:
            fig_sales = px.line(
                monthly_summary,
                x="Order Month",
                y="总销售额",
                markers=True,
                title="各月总销售额 (Total Cost)",
                color_discrete_sequence=["#28B463"],
            )
            fig_sales.update_layout(xaxis_title="月份", yaxis_title="销售额")
            st.plotly_chart(fig_sales, use_container_width=True)

        with st.expander("查看各月汇总数据表", expanded=True):
            st.dataframe(
                monthly_summary.style.format({"总销量": "{:,.0f}", "总销售额": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ==========================================
        # 模块二：Vendor SKU 每月销量和销售额看板
        # ==========================================
        st.subheader("📦 2. Vendor SKU 每月销量与销售额分析")

        # 聚合 SKU 月度数据（带上 Description 方便识别产品）
        group_cols = ["Vendor SKU", "Order Month"]
        if "Description" in filtered_df.columns:
            # 取每个 SKU 对应的第一个非空描述
            sku_desc_map = filtered_df.groupby("Vendor SKU")["Description"].first().to_dict()
        else:
            sku_desc_map = {}

        sku_monthly_df = (
            filtered_df.groupby(["Vendor SKU", "Order Month"], as_index=False)
            .agg(销量=("Quantity", "sum"), 销售额=("Total Cost", "sum"))
            .sort_values(["Vendor SKU", "Order Month"])
        )
        if sku_desc_map:
            sku_monthly_df.insert(1, "Description", sku_monthly_df["Vendor SKU"].map(sku_desc_map))

        # 生成横向透视矩阵（行：SKU，列：月份）
        sku_qty_pivot = filtered_df.pivot_table(
            index="Vendor SKU",
            columns="Order Month",
            values="Quantity",
            aggfunc="sum",
            fill_value=0,
        )
        sku_qty_pivot["合计销量"] = sku_qty_pivot.sum(axis=1)
        sku_qty_pivot = sku_qty_pivot.sort_values("合计销量", ascending=False)

        sku_sales_pivot = filtered_df.pivot_table(
            index="Vendor SKU",
            columns="Order Month",
            values="Total Cost",
            aggfunc="sum",
            fill_value=0,
        )
        sku_sales_pivot["合计销售额"] = sku_sales_pivot.sum(axis=1)
        sku_sales_pivot = sku_sales_pivot.sort_values("合计销售额", ascending=False)

        # Top N SKU 趋势对比图
        top_n = st.slider("选择展示 Top N 畅销 Vendor SKU 趋势对比", min_value=3, max_value=20, value=5)
        top_skus = sku_sales_pivot.head(top_n).index.tolist()
        top_sku_trend = sku_monthly_df[sku_monthly_df["Vendor SKU"].isin(top_skus)]

        tab_chart1, tab_chart2 = st.tabs(["Top SKU 月度销售额对比", "Top SKU 月度销量对比"])
        with tab_chart1:
            fig_sku_sales = px.bar(
                top_sku_trend,
                x="Order Month",
                y="销售额",
                color="Vendor SKU",
                barmode="group",
                title=f"Top {top_n} Vendor SKU 每月销售额对比",
            )
            st.plotly_chart(fig_sku_sales, use_container_width=True)

        with tab_chart2:
            fig_sku_qty = px.bar(
                top_sku_trend,
                x="Order Month",
                y="销量",
                color="Vendor SKU",
                barmode="group",
                title=f"Top {top_n} Vendor SKU 每月销量对比",
            )
            st.plotly_chart(fig_sku_qty, use_container_width=True)

        # 交互式筛选特定 SKU 查看明细
        selected_skus = st.multiselect(
            "搜索或选择特定 Vendor SKU 查看详细矩阵（留空则显示全部）",
            options=sku_sales_pivot.index.tolist(),
        )

        display_qty_pivot = sku_qty_pivot.loc[selected_skus] if selected_skus else sku_qty_pivot
        display_sales_pivot = sku_sales_pivot.loc[selected_skus] if selected_skus else sku_sales_pivot
        display_detail_df = (
            sku_monthly_df[sku_monthly_df["Vendor SKU"].isin(selected_skus)]
            if selected_skus
            else sku_monthly_df
        )

        tab1, tab2, tab3 = st.tabs(["📊 SKU 各月销量矩阵", "💰 SKU 各月销售额矩阵", "📋 SKU 月度明细列表"])

        with tab1:
            st.dataframe(display_qty_pivot.style.format("{:,.0f}"), use_container_width=True)

        with tab2:
            st.dataframe(display_sales_pivot.style.format("{:,.2f}"), use_container_width=True)

        with tab3:
            st.dataframe(
                display_detail_df.style.format({"销量": "{:,.0f}", "销售额": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

        # ==========================================
        # 模块三：一键下载完整 Excel 报表
        # ==========================================
        st.divider()
        st.subheader("📥 3. 下载分析报表")

        excel_data = convert_df_to_excel(
            monthly_summary,
            display_qty_pivot,
            display_sales_pivot,
            display_detail_df,
        )

        st.download_button(
            label="⬇️ 下载完整 Excel 报表 (含各月汇总 & SKU 矩阵)",
            data=excel_data,
            file_name="PO_Sales_and_SKU_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

else:
    st.info("👈 请在左侧边栏上传包含订单表头的 Excel 或 CSV 文件以生成看板。")
