import pandas as pd
import streamlit as st

# --------------------------
# 1. 读取Excel数据（适配云端+无转义错误）
# --------------------------
def get_dataframe_from_excel():
    try:
        # 核心：仅保留相对路径，彻底删除本地C盘路径引用
        excel_file_path = "（商场销售数据）supermarket_sales.xlsx"
        
        # 读取Excel文件
        df = pd.read_excel(
            excel_file_path,
            sheet_name='销售数据',
            skiprows=1,
            index_col='订单号',
            engine='openpyxl'
        )
        
        # 数据预处理
        df['小时数'] = pd.to_datetime(df["时间"], format="%H:%M:%S").dt.hour
        
        return df

    # 异常处理（移除所有含\的本地路径提示，避免转义错误）
    except FileNotFoundError:
        st.error(r"""❌ 文件未找到！请检查：
        1. Excel文件是否上传到项目根目录
        2. 文件名是否为：（商场销售数据）supermarket_sales.xlsx
        3. 请勿使用本地电脑路径，仅保留相对路径""")
        return pd.DataFrame()
    
    except ImportError:
        st.error("❌ 缺少openpyxl依赖！请检查requirements.txt是否包含openpyxl")
        return pd.DataFrame()
    
    except ValueError as e:
        st.error(f"❌ 数据格式错误：{e}\n请检查Excel是否有'销售数据'工作表、'订单号'/'时间'列")
        return pd.DataFrame()
    
    except Exception as e:
        st.error(f"❌ 未知错误：{str(e)}")
        return pd.DataFrame()

# --------------------------
# 2. 侧边栏筛选功能
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
    
    # 应用筛选条件
    df_filtered = df.query(
        "城市 == @city_selected & 顾客类型 == @customer_selected & 性别 == @gender_selected"
    )
    return df_filtered

# --------------------------
# 3. 主程序入口
# --------------------------
if __name__ == "__main__":
    # 页面配置
    st.set_page_config(
        page_title="商场销售数据筛选工具",
        page_icon="📊",
        layout="wide"
    )
    
    # 页面标题
    st.title("📊 商场销售数据筛选分析工具")
    st.divider()
    
    # 读取数据
    sale_df = get_dataframe_from_excel()
    
    # 展示结果
    if not sale_df.empty:
        df_final = add_sidebar_func(sale_df)
        st.subheader("📋 筛选后的数据结果")
        st.dataframe(df_final, use_container_width=True)
        
        # 统计信息
        st.info(f"""✅ 筛选结果统计：
        - 总行数：{df_final.shape[0]} 行
        - 总列数：{df_final.shape[1]} 列
        - 涉及城市：{', '.join(df_final['城市'].unique())}""")
    else:
        st.warning("⚠️ 暂无数据可展示，请检查Excel文件和依赖配置！")
