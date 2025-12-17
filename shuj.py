import pandas as pd
import streamlit as st
import plotly.express as px  # 用于绘制交互式条形图

# --------------------------
# 1. 读取Excel数据（适配云端+无转义错误）
# --------------------------
def get_dataframe_from_excel():
    try:
        # 仅保留相对路径，适配云端
        excel_file_path = "（商场销售数据）supermarket_sales.xlsx"
        
        df = pd.read_excel(
            excel_file_path,
            sheet_name='销售数据',
            skiprows=1,
            index_col='订单号',
            engine='openpyxl'
        )
        
        # 数据预处理：提取小时数+计算销售额（若Excel无销售额列则自动计算）
        df['小时数'] = pd.to_datetime(df["时间"], format="%H:%M:%S").dt.hour
        if "销售额" not in df.columns and "单价" in df.columns and "数量" in df.columns:
            df['销售额'] = df["单价"] * df["数量"]
        
        return df

    except FileNotFoundError:
        st.error(r"""❌ 文件未找到！请检查：
        1. Excel文件是否上传到项目根目录
        2. 文件名是否为：（商场销售数据）supermarket_sales.xlsx
        3. 请勿使用本地电脑路径，仅保留相对路径""")
        return pd.DataFrame()
    
    except ImportError:
        st.error("❌ 缺少依赖库！请检查requirements.txt包含：pandas、streamlit、openpyxl、plotly")
        return pd.DataFrame()
    
    except ValueError as e:
        st.error(f"❌ 数据格式错误：{e}\n请检查Excel是否有'销售数据'工作表、'订单号'/'时间'列")
        return pd.DataFrame()
    
    except Exception as e:
        st.error(f"❌ 未知错误：{str(e)}")
        return pd.DataFrame()

# --------------------------
# 2. 侧边栏筛选功能（增强版）
# --------------------------
def add_sidebar_func(df):
    with st.sidebar:
        st.header("🔍 数据筛选条件")
        
        # 城市筛选
        city_options = df["城市"].unique()
        city_selected = st.multiselect(
            "选择城市",
            options=city_options,
            default=city_options
        )
        
        # 顾客类型筛选
        customer_options = df["顾客类型"].unique()
        customer_selected = st.multiselect(
            "选择顾客类型",
            options=customer_options,
            default=customer_options
        )
        
        # 性别筛选
        gender_options = df["性别"].unique()
        gender_selected = st.multiselect(
            "选择性别",
            options=gender_options,
            default=gender_options
        )

        # 新增：时段筛选（按小时）
        st.subheader("⏰ 时段筛选")
        hour_min = int(df["小时数"].min())
        hour_max = int(df["小时数"].max())
        hour_range = st.slider(
            "选择交易小时范围",
            min_value=hour_min,
            max_value=hour_max,
            value=(hour_min, hour_max)
        )
    
    # 应用所有筛选条件
    df_filtered = df.query(
        "城市 == @city_selected & 顾客类型 == @customer_selected & 性别 == @gender_selected & 小时数 >= @hour_range[0] & 小时数 <= @hour_range[1]"
    )
    return df_filtered

# --------------------------
# 3. 条形图绘制函数（核心优化）
# --------------------------
def plot_bar_charts(df):
    st.subheader("📈 数据可视化 - 条形图分析")
    
    # 分栏展示多个条形图
    col1, col2 = st.columns(2)
    
    # 图1：各城市销售额对比
    with col1:
        city_sales = df.groupby("城市")["销售额"].sum().reset_index()
        fig1 = px.bar(
            city_sales,
            x="城市",
            y="销售额",
            title="各城市销售额对比",
            color="城市",
            color_discrete_sequence=px.colors.qualitative.Set2,
            labels={"销售额": "销售额（元）", "城市": "城市名称"},
            height=400
        )
        # 优化图表样式
        fig1.update_layout(
            title_font_size=14,
            xaxis_title_font_size=12,
            yaxis_title_font_size=12,
            showlegend=False
        )
        st.plotly_chart(fig1, use_container_width=True)
    
    # 图2：各时段销量对比
    with col2:
        hour_sales = df.groupby("小时数")["数量"].sum().reset_index()
        fig2 = px.bar(
            hour_sales,
            x="小时数",
            y="数量",
            title="各时段销量对比",
            color="数量",
            color_continuous_scale="Blues",
            labels={"数量": "销量（件）", "小时数": "交易小时（点）"},
            height=400
        )
        fig2.update_layout(
            title_font_size=14,
            xaxis_title_font_size=12,
            yaxis_title_font_size=12,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig2, use_container_width=True)
    
    # 图3：性别+品类销售额分布（新增）
    st.subheader("🎯 性别&品类销售额分析")
    # 确保"商品类别"列存在（适配常见销售数据字段）
    if "商品类别" in df.columns:
        gender_category = df.groupby(["性别", "商品类别"])["销售额"].sum().reset_index()
        fig3 = px.bar(
            gender_category,
            x="商品类别",
            y="销售额",
            color="性别",
            barmode="group",  # 分组条形图
            title="不同性别顾客的品类消费对比",
            labels={"销售额": "销售额（元）", "商品类别": "商品品类", "性别": "顾客性别"},
            height=450
        )
        fig3.update_layout(title_font_size=14)
        st.plotly_chart(fig3, use_container_width=True)

# --------------------------
# 4. 主程序入口
# --------------------------
if __name__ == "__main__":
    # 页面配置（宽屏适配可视化）
    st.set_page_config(
        page_title="商场销售数据可视化工具",
        page_icon="📊",
        layout="wide"
    )
    
    # 页面标题+说明
    st.title("📊 商场销售数据可视化分析工具")
    st.markdown("### 基于条形图的多维度销售数据分析（筛选+可视化）")
    st.divider()
    
    # 读取数据
    sale_df = get_dataframe_from_excel()
    
    # 数据非空时展示筛选+可视化
    if not sale_df.empty:
        df_final = add_sidebar_func(sale_df)
        
        # 展示筛选后的数据概览
        st.subheader("📋 筛选后数据概览")
        col_info1, col_info2, col_info3 = st.columns(3)
        with col_info1:
            st.metric("总数据行数", df_final.shape[0])
        with col_info2:
            st.metric("总销售额", f"{df_final['销售额'].sum():,.2f} 元")
        with col_info3:
            st.metric("总销量", f"{df_final['数量'].sum()} 件")
        
        # 可选展示原始数据（折叠面板）
        with st.expander("点击查看筛选后原始数据"):
            st.dataframe(df_final, use_container_width=True)
        
        # 绘制条形图
        plot_bar_charts(df_final)
    
    else:
        st.warning("⚠️ 暂无数据可展示，请检查Excel文件和依赖配置！")
