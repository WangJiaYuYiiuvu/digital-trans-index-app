import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO
import os

# 设置中文字体支持
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

# 页面配置
st.set_page_config(
    page_title="上市公司数字化转型指数查询系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 实际Excel文件路径
DEFAULT_FILE_PATH = r"C:\Users\Lenovo\Desktop\text\数字化转型词频统计结果（总）.xlsx"  

# 行业映射配置
INDUSTRY_MAPPING = {
    '金融': ['银行', '保险', '证券', '金融', '基金', '投资', '信托'],
    '房地产': ['地产', '置业', '房产', '物业', '房地产'],
    '制造业': ['制造', '工业', '科技', '电子', '机械', '设备', '汽车'],
    '交通运输': ['航空', '铁路', '物流', '港口', '运输', '交通', '航运'],
    '能源': ['电力', '石油', '煤炭', '能源', '燃气', '新能源'],
    '信息技术': ['信息', '技术', '软件', '互联网', '计算机', '通信'],
    '医疗健康': ['医疗', '医药', '健康', '生物', '制药'],
    '消费': ['零售', '食品', '饮料', '消费', '家电', '服装'],
    '教育': ['教育', '培训', '学校'],
    '传媒': ['传媒', '广告', '娱乐', '影视', '出版']
}

# 指数计算权重（可调整）
TECH_WEIGHTS = {
    '人工智能技术': 0.3,
    '区块链技术': 0.1,
    '大数据技术': 0.25,
    '云计算技术': 0.15,
    '数字技术应用': 0.2
}

@st.cache_data(show_spinner="正在加载并计算指数...")
def load_and_calculate_index():
    """加载数据并计算数字化转型指数"""
    # 检查文件是否存在
    if not os.path.exists(DEFAULT_FILE_PATH):
        st.error(f"❌ 文件不存在: {DEFAULT_FILE_PATH}")
        return pd.DataFrame()
    
    try:
        # 读取原始词频数据
        df = pd.read_excel(DEFAULT_FILE_PATH)
        
        # 计算数字化转型指数（加权求和）
        df['数字化转型指数'] = df.apply(
            lambda row: sum(row[tech] * weight for tech, weight in TECH_WEIGHTS.items()),
            axis=1
        )
        
        # 匹配行业（基于企业名称）
        df['行业'] = '其他'
        for industry, keywords in INDUSTRY_MAPPING.items():
            pattern = '|'.join(keywords)
            df.loc[df['企业名称'].str.contains(pattern, case=False), '行业'] = industry
        
        st.success(f"✅ 成功加载数据，共 {len(df)} 家企业（已计算指数和行业）")
        return df
    except Exception as e:
        st.error(f"❌ 数据处理失败: {str(e)}")
        return pd.DataFrame()

def visualize_index_distribution(df):
    """可视化数字化转型指数分布"""
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.hist(df['数字化转型指数'], bins=15, alpha=0.7, color='teal')
    
    ax.set_xlabel('数字化转型指数')
    ax.set_ylabel('企业数量')
    ax.set_title('上市公司数字化转型指数分布')
    
    plt.tight_layout()
    return fig

def main():
    st.title("上市公司数字化转型指数查询系统")
    st.markdown("""
    本系统基于**大数据文本挖掘**技术，通过年报词频统计计算数字化转型指数，支持多维度筛选查询。
    """)
    
    # 加载并计算指数
    df = load_and_calculate_index()
    
    # 检查数据是否有效
    if df.empty:
        st.warning("数据加载失败，请检查文件路径或格式")
        return
    
    # 侧边栏筛选器
    st.sidebar.header("筛选条件")
    
    # 行业筛选
    industries = df['行业'].unique().tolist()
    selected_industries = st.sidebar.multiselect("选择行业", industries, default=industries)
    
    # 指数范围筛选
    min_index = df['数字化转型指数'].min()
    max_index = df['数字化转型指数'].max()
    index_range = st.sidebar.slider(
        "数字化转型指数范围",
        min_value=min_index,
        max_value=max_index,
        value=(min_index, max_index),
        format="%.2f"
    )
    
    # 股票代码/企业名称搜索
    search_term = st.sidebar.text_input("搜索股票代码/企业名称", "")
    
    # 应用筛选
    filtered_df = df.copy()
    
    # 行业筛选
    if selected_industries:
        filtered_df = filtered_df[filtered_df['行业'].isin(selected_industries)]
    
    # 指数范围筛选
    filtered_df = filtered_df[
        (filtered_df['数字化转型指数'] >= index_range[0]) & 
        (filtered_df['数字化转型指数'] <= index_range[1])
    ]
    
    # 关键词搜索
    if search_term:
        filtered_df = filtered_df[
            filtered_df['股票代码'].str.contains(search_term, case=False) | 
            filtered_df['企业名称'].str.contains(search_term, case=False)
        ]
    
    # 展示筛选结果
    st.subheader(f"查询结果（共 {len(filtered_df)} 家企业）")
    
    # 数据表格展示
    st.dataframe(
        filtered_df[['股票代码', '企业名称', '行业', '数字化转型指数', '总词频数']],
        use_container_width=True,
        hide_index=True,
        column_config={
            "数字化转型指数": st.column_config.NumberColumn(format="%.2f"),
            "总词频数": st.column_config.NumberColumn()
        }
    )
    
    # 指数分布可视化
    st.subheader("数字化转型指数分布")
    fig = visualize_index_distribution(filtered_df)
    st.pyplot(fig)
    
    # 指数Top10企业
    st.subheader("数字化转型指数Top10企业")
    top10_df = filtered_df.sort_values(by='数字化转型指数', ascending=False).head(10)
    st.dataframe(
        top10_df[['股票代码', '企业名称', '行业', '数字化转型指数']],
        use_container_width=True,
        hide_index=True
    )

if __name__ == "__main__":
    main()