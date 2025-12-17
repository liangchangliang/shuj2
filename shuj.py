import pandas as pd
import streamlit as st

# 读取Excel数据的函数（固定路径+C:\Users\712\Desktop + 显式指定openpyxl引擎）
def get_dataframe_from_excel():
    try:
        # 固定Excel文件路径（C:\Users\712\Desktop），原始字符串避免转义
        excel_file_path = r"C:\Users\712\Desktop\（商场销售数据）supermarket_sales.xlsx"
        
        # 读取Excel，显式指定openpyxl引擎（确保依赖生效）
        df = pd.read_excel(
            excel_file_path,
            sheet_name='销售数据',
            skiprows=1,
            index_col='订单号',
            engine='openpyxl'  # 强制使用openpyxl解析xlsx文件
        )
        
        # 提取交易小时数（数据预处理）
        df['小时数'] = pd.to_datetime(df["时间"], format="%H:%M:%S").dt.hour
        return df
    
    # 针对性异常处理（方便排查问题）
    except ImportError:
        st.error("❌ 缺少openpyxl库！请打开CMD执行：pip install openpyxl")
        return pd.DataFrame()
    except FileNotFoundError:
        st.error(f"❌ 未找到Excel文件！请确认文件路径：\n{excel_file_path}")
        return pd.DataFrame()
    except ValueError as e:
        st.error(f"❌ Excel数据格式错误：{str(e)}\n请检查'销售数据'工作表是否有'订单号'/'时间'等列")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ 读取Excel出错：{str(e)}")
        return pd.DataFrame()

# 侧边栏筛选功能（无修改）
def add_sidebar_func(df):
    with st.sidebar:
        st.header("🔍 数据筛选")
        # 城市筛选
        city_unique = df["城市"].unique()
        city = st.multiselect("选择城市：", options=city_unique, default=city_unique)
        
        # 顾客类型筛选
        customer_type_unique = df["顾客类型"].unique()
        customer_type = st.multiselect("选择顾客类型：", options=customer_type_unique, default=customer_type_unique)
        
        # 性别筛选
        gender_unique = df["性别"].unique()
        gender = st.multiselect("选择性别：", options=gender_unique, default=gender_unique)
    
    # 应用筛选条件
    df_selection = df.query("城市 == @city & 顾客类型 ==@customer_type & 性别 == @gender")
    return df_selection

# 主程序入口
if __name__ == "__main__":
    # 设置页面标题
    st.set_page_config(page_title="商场销售数据筛选", page_icon="📊")
    st.title("📊 商场销售数据筛选工具")
    st.divider()
    
    # 读取Excel数据
    sale_df = get_dataframe_from_excel()
    
    # 数据非空时展示筛选结果
    if not sale_df.empty:
        df_filtered = add_sidebar_func(sale_df)
        
        # 展示筛选后的数据
        st.subheader("筛选后的数据")
        st.dataframe(df_filtered, use_container_width=True)
        st.info(f"✅ 筛选后数据共 {df_filtered.shape[0]} 行，{df_filtered.shape[1]} 列")
    else:
        st.warning("⚠️ 暂无数据可展示，请检查Excel文件或依赖库！")
