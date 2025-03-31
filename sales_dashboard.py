import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import traceback
import os
import json
from pathlib import Path

# 设置页面配置
st.set_page_config(
    page_title="销售数据分析仪表盘",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 定义配置文件路径
CONFIG_PATH = "./.streamlit/dashboard_config.json"


# ---- 配置加载与保存函数 ----
def load_config():
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            # 确保.streamlit目录存在
            os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
            # 创建默认配置
            default_config = {
                "default_file_path": "C:/Users/何晴雅/Desktop/Q1xlsx.xlsx",
                "tableau_theme": True,
                "last_uploaded_file": None
            }
            save_config(default_config)
            return default_config
    except Exception as e:
        st.error(f"加载配置文件时出错: {str(e)}")
        return {
            "default_file_path": "C:/Users/何晴雅/Desktop/Q1xlsx.xlsx",
            "tableau_theme": True,
            "last_uploaded_file": None
        }


def save_config(config):
    try:
        # 确保.streamlit目录存在
        os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"保存配置文件时出错: {str(e)}")


# 加载配置
if 'config' not in st.session_state:
    st.session_state.config = load_config()

# 初始化session state变量
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'file_path' not in st.session_state:
    st.session_state.file_path = st.session_state.config['default_file_path']
if 'is_sample_data' not in st.session_state:
    st.session_state.is_sample_data = True

# 定义一些更美观的Tableau风格CSS样式
st.markdown("""
<style>
    /* === Tableau风格主题 === */

    /* 主要配色 */
    :root {
        --tableau-blue: #4E79A7;
        --tableau-orange: #F28E2B;
        --tableau-red: #E15759;
        --tableau-teal: #59A14F;
        --tableau-green: #76B7B2;
        --tableau-yellow: #EDC948;
        --tableau-purple: #B07AA1;
        --tableau-pink: #FF9DA7;
        --tableau-brown: #9C755F;
        --tableau-gray: #BAB0AC;
        --tableau-light-bg: #F5F5F5;
        --tableau-dark-text: #333333;
        --tableau-medium-text: #666666;
        --tableau-light-border: #E0E0E0;
    }

    /* 整体背景和字体 */
    body {
        background-color: var(--tableau-light-bg);
        font-family: 'Segoe UI', 'Arial', sans-serif;
        color: var(--tableau-dark-text);
    }

    /* 主标题 */
    .main-header {
        font-size: 2.8rem;
        color: var(--tableau-dark-text);
        text-align: center;
        margin-bottom: 2rem;
        padding: 1.5rem;
        background-color: white;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        font-weight: 600;
        letter-spacing: -0.5px;
    }

    /* 次级标题 */
    .sub-header {
        font-size: 1.8rem;
        color: var(--tableau-dark-text);
        padding-top: 1.5rem;
        padding-bottom: 1rem;
        margin-top: 1rem;
        border-bottom: 2px solid var(--tableau-light-border);
        font-weight: 500;
        letter-spacing: -0.3px;
    }

    /* 卡片容器 */
    .card {
        border-radius: 8px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        background-color: white;
        transition: transform 0.3s, box-shadow 0.3s;
        border: 1px solid var(--tableau-light-border);
    }

    .card:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.1);
    }

    /* 指标值样式 */
    .metric-value {
        font-size: 2.2rem;
        font-weight: 600;
        color: var(--tableau-blue);
        margin: 0.5rem 0;
    }

    .metric-label {
        font-size: 1.1rem;
        color: var(--tableau-medium-text);
        font-weight: 500;
    }

    /* 高亮内容 */
    .highlight {
        background-color: rgba(78, 121, 167, 0.1);
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1.5rem 0;
        border-left: 5px solid var(--tableau-blue);
    }

    /* 选项卡样式 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }

    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        border-radius: 5px 5px 0 0;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background-color: rgba(78, 121, 167, 0.1);
        border-bottom: 3px solid var(--tableau-blue);
    }

    /* 折叠面板 */
    .stExpander {
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        margin-bottom: 1rem;
    }

    /* 下载按钮 */
    .download-button {
        text-align: center;
        margin-top: 2rem;
    }

    /* 章节间距 */
    .section-gap {
        margin-top: 2.5rem;
        margin-bottom: 2.5rem;
    }

    /* 调整图表容器样式 */
    .st-emotion-cache-1wrcr25 {
        margin-top: 2rem !important;
        margin-bottom: 3rem !important;
        padding: 1rem !important;
        background-color: white !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05) !important;
    }

    /* 设置侧边栏样式 */
    .st-emotion-cache-6qob1r {
        background-color: white;
        border-right: 1px solid var(--tableau-light-border);
    }

    [data-testid="stSidebar"] {
        background-color: white;
        padding: 1rem;
    }

    [data-testid="stSidebarNav"] {
        padding-top: 2rem;
    }

    .sidebar-header {
        font-size: 1.3rem;
        color: var(--tableau-dark-text);
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid var(--tableau-light-border);
        font-weight: 500;
    }

    /* 调整图表字体大小 */
    .js-plotly-plot .plotly .ytick text, 
    .js-plotly-plot .plotly .xtick text {
        font-size: 14px !important;
    }

    .js-plotly-plot .plotly .gtitle {
        font-size: 18px !important;
    }

    /* 错误消息样式 */
    .error-message {
        color: #721c24;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }

    /* 信息消息样式 */
    .info-message {
        color: #0c5460;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }

    /* 按钮样式优化 */
    .stButton > button {
        background-color: var(--tableau-blue);
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 4px;
        font-weight: 500;
        transition: all 0.3s;
    }

    .stButton > button:hover {
        background-color: #3d6285; 
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }

    /* 图表容器样式 */
    .chart-container {
        background-color: white;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        margin-bottom: 2rem;
    }

    /* 表格样式 */
    .dataframe {
        border: 1px solid var(--tableau-light-border);
        border-radius: 8px;
        overflow: hidden;
    }

    .dataframe th {
        background-color: rgba(78, 121, 167, 0.1);
        color: var(--tableau-dark-text);
        font-weight: 500;
        padding: 0.75rem 1rem !important;
        border-bottom: 1px solid var(--tableau-light-border);
    }

    .dataframe td {
        padding: 0.75rem 1rem !important;
        border-bottom: 1px solid var(--tableau-light-border);
        background-color: white;
    }

    /* 过滤器样式 */
    .filter-container {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }

    /* 文件上传区域 */
    .upload-container {
        background-color: white;
        padding: 1.5rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        border: 2px dashed var(--tableau-light-border);
    }

    /* 状态指示图标 */
    .status-icon {
        font-size: 1.2rem;
        margin-right: 0.5rem;
    }

    /* 成功状态 */
    .success-status {
        color: var(--tableau-teal);
    }

    /* 警告状态 */
    .warning-status {
        color: var(--tableau-orange);
    }

    /* 错误状态 */
    .error-status {
        color: var(--tableau-red);
    }

    /* 默认文件配置容器 */
    .config-container {
        background-color: rgba(78, 121, 167, 0.05);
        padding: 1rem;
        border-radius: 8px;
        margin-top: 1rem;
        border: 1px solid rgba(78, 121, 167, 0.2);
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown('<div class="main-header">销售数据分析仪表盘</div>', unsafe_allow_html=True)


# 格式化数值的函数
def format_yuan(value):
    if value >= 10000:
        return f"{value / 10000:.1f}万元"
    return f"{value:,.0f}元"


# 安全筛选函数
def safe_filter(df, filter_func):
    """安全应用筛选条件，确保筛选后的数据不为空"""
    try:
        filtered = filter_func(df)
        if filtered.empty:
            st.warning("当前筛选条件下没有匹配的数据。请尝试放宽筛选条件。")
            return df  # 返回原始数据而不是空数据
        return filtered
    except Exception as e:
        st.error(f"筛选数据时出错: {str(e)}")
        return df


# 加载数据函数 - 修改以支持默认文件路径和session state
@st.cache_data
def load_data(file_path=None):
    # 如果提供了文件路径，从文件加载
    try:
        if file_path is not None:
            try:
                # 检查是否是FileUploader对象还是字符串路径
                if hasattr(file_path, 'read'):
                    # 是上传的文件对象
                    df = pd.read_excel(file_path, engine='openpyxl')
                    # 更新配置中的最后一次上传路径
                    st.session_state.config["last_uploaded_file"] = file_path.name
                    save_config(st.session_state.config)
                else:
                    # 是文件路径字符串
                    df = pd.read_excel(file_path, engine='openpyxl')
            except Exception as e:
                st.error(f"文件加载失败: {str(e)}。使用示例数据进行演示。")
                df = load_sample_data()
                return df, True  # 返回示例数据标记
        else:
            df = load_sample_data()
            return df, True  # 返回示例数据标记

        # 数据预处理
        df['销售额'] = df['单价（箱）'] * df['数量（箱）']

        # 确保发运月份是日期类型
        try:
            df['发运月份'] = pd.to_datetime(df['发运月份'])
        except Exception as e:
            st.info(f"发运月份转换为日期类型时出错。原因：{str(e)}。将保持原格式。")

        # 添加简化产品名称列
        df['简化产品名称'] = df.apply(lambda row: get_simplified_product_name(row['产品代码'], row['产品名称']), axis=1)

        return df, False  # 返回实际数据标记
    except Exception as e:
        st.error(f"加载数据时出现未预期的错误: {str(e)}")
        st.write("错误详情:")
        st.write(traceback.format_exc())
        df = load_sample_data()
        return df, True  # 返回示例数据标记


# 创建产品代码到简化产品名称的映射函数 (修复版)
def get_simplified_product_name(product_code, product_name):
    try:
        # 从产品名称中提取关键部分
        if '口力' in product_name:
            # 提取"口力"之后的产品类型
            name_parts = product_name.split('口力')[1].split('-')[0].strip()
            # 进一步简化，只保留主要部分（去掉规格和包装形式）
            for suffix in ['G分享装袋装', 'G盒装', 'G袋装', 'KG迷你包', 'KG随手包']:
                name_parts = name_parts.split(suffix)[0]

            # 去掉可能的数字和单位
            import re
            simple_name = re.sub(r'\d+\w*\s*', '', name_parts).strip()

            # 始终包含产品代码以确保唯一性
            return f"{simple_name} ({product_code})"
        else:
            # 如果无法提取，则返回产品代码
            return product_code
    except Exception as e:
        # 出错时返回产品代码
        return product_code


# 创建示例数据（以防用户没有上传文件）
@st.cache_data
def load_sample_data():
    # 创建简化版示例数据
    data = {
        '客户简称': ['广州佳成行', '广州佳成行', '广州佳成行', '广州佳成行', '广州佳成行',
                     '广州佳成行', '河南甜丰號', '河南甜丰號', '河南甜丰號', '河南甜丰號',
                     '河南甜丰號', '广州佳成行', '河南甜丰號', '广州佳成行', '河南甜丰號',
                     '广州佳成行'],
        '所属区域': ['南', '南', '南', '南', '南', '南', '中', '中', '中', '中', '中',
                     '南', '中', '南', '中', '南'],
        '发运月份': ['2025-03', '2025-03', '2025-03', '2025-03', '2025-03', '2025-03',
                     '2025-03', '2025-03', '2025-03', '2025-03', '2025-03', '2025-03',
                     '2025-03', '2025-03', '2025-03', '2025-03'],
        '申请人': ['梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽', '梁洪泽',
                   '胡斌', '胡斌', '胡斌', '胡斌', '胡斌', '梁洪泽', '胡斌', '梁洪泽',
                   '胡斌', '梁洪泽'],
        '产品代码': ['F3415D', 'F3421D', 'F0104J', 'F0104L', 'F3411A', 'F01E4B',
                     'F01L4C', 'F01C2P', 'F01E6D', 'F3450B', 'F3415B', 'F0110C',
                     'F0183F', 'F01K8A', 'F0183K', 'F0101P'],
        '产品名称': ['口力酸小虫250G分享装袋装-中国', '口力可乐瓶250G分享装袋装-中国',
                     '口力比萨XXL45G盒装-中国', '口力比萨68G袋装-中国', '口力午餐袋77G袋装-中国',
                     '口力汉堡108G袋装-中国', '口力扭扭虫2KG迷你包-中国', '口力字节软糖2KG迷你包-中国',
                     '口力西瓜1.5KG随手包-中国', '口力七彩熊1.5KG随手包-中国', '口力酸小虫1.5KG随手包-中国',
                     '口力软糖新品A-中国', '口力软糖新品B-中国', '口力软糖新品C-中国', '口力软糖新品D-中国',
                     '口力软糖新品E-中国'],
        '订单类型': ['订单-正常产品'] * 16,
        '单价（箱）': [121.44, 121.44, 216.96, 126.72, 137.04, 137.04, 127.2, 127.2,
                     180, 180, 180, 150, 160, 170, 180, 190],
        '数量（箱）': [10, 10, 20, 50, 252, 204, 7, 2, 6, 6, 6, 30, 20, 15, 10, 5]
    }

    df = pd.DataFrame(data)
    df['销售额'] = df['单价（箱）'] * df['数量（箱）']

    # 添加简化产品名称列
    df['简化产品名称'] = df.apply(lambda row: get_simplified_product_name(row['产品代码'], row['产品名称']), axis=1)

    return df


# 侧边栏 - 配置和上传
st.sidebar.markdown('<div class="sidebar-header">数据配置</div>', unsafe_allow_html=True)

# 文件上传前的说明
st.sidebar.markdown('<div class="upload-container">', unsafe_allow_html=True)
st.sidebar.markdown("""
### 使用说明
* 上传Excel表格数据进行分析
* 分享链接后，他人将看到您上传的数据
* 默认文件路径已设置为您的文件位置
""")
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# 文件上传组件
uploaded_file = st.sidebar.file_uploader("上传Excel销售数据文件", type=["xlsx", "xls"])

# 默认文件路径配置
with st.sidebar.expander("默认文件设置", expanded=False):
    default_file = st.text_input(
        "设置默认文件路径",
        value=st.session_state.config["default_file_path"],
        help="设置默认加载的Excel文件路径，分享链接后将自动加载此文件"
    )

    if st.button("保存默认文件设置"):
        st.session_state.config["default_file_path"] = default_file
        save_config(st.session_state.config)
        st.success("默认文件路径已保存！")
        st.session_state.file_path = default_file
        # 清除之前缓存的数据，使其重新加载
        st.session_state.data_loaded = False
        st.experimental_rerun()

# 加载数据逻辑 - 优先使用上传的文件，其次使用默认路径
try:
    # 检查是否需要加载数据（未加载或重新上传）
    if not st.session_state.data_loaded or uploaded_file is not None:
        if uploaded_file is not None:
            # 用户刚刚上传了新文件
            df, is_sample = load_data(uploaded_file)
            st.session_state.df = df
            st.session_state.is_sample_data = is_sample
            st.session_state.data_loaded = True

            if not is_sample:
                st.sidebar.success(f"""
                <div class="success-status">
                    <span class="status-icon">✅</span> 已成功加载文件: {uploaded_file.name}
                </div>
                """, unsafe_allow_html=True)

        elif not st.session_state.data_loaded:
            # 尝试从默认路径加载
            try:
                default_path = st.session_state.config["default_file_path"]
                if os.path.exists(default_path):
                    df, is_sample = load_data(default_path)
                    st.session_state.df = df
                    st.session_state.is_sample_data = is_sample
                    st.session_state.data_loaded = True

                    if not is_sample:
                        st.sidebar.success(f"""
                        <div class="success-status">
                            <span class="status-icon">✅</span> 已从默认路径加载文件: {os.path.basename(default_path)}
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    # 默认文件不存在，使用示例数据
                    df = load_sample_data()
                    st.session_state.df = df
                    st.session_state.is_sample_data = True
                    st.session_state.data_loaded = True

                    st.sidebar.warning(f"""
                    <div class="warning-status">
                        <span class="status-icon">⚠️</span> 默认文件不存在，使用示例数据。请上传您的文件。
                    </div>
                    """, unsafe_allow_html=True)
            except Exception as e:
                # 出错则使用示例数据
                df = load_sample_data()
                st.session_state.df = df
                st.session_state.is_sample_data = True
                st.session_state.data_loaded = True

                st.sidebar.error(f"""
                <div class="error-status">
                    <span class="status-icon">❌</span> 加载默认文件出错: {str(e)}。使用示例数据。
                </div>
                """, unsafe_allow_html=True)
    else:
        # 使用已加载的数据
        df = st.session_state.df

        if st.session_state.is_sample_data:
            st.sidebar.info("""
            <div class="info-message">
                <span class="status-icon">ℹ️</span> 正在使用示例数据。请上传您的数据文件获取真实分析。
            </div>
            """, unsafe_allow_html=True)

except Exception as e:
    st.error(f"加载数据时出错: {str(e)}")
    df = load_sample_data()
    st.session_state.df = df
    st.session_state.is_sample_data = True
    st.session_state.data_loaded = True

    st.sidebar.warning("""
    <div class="warning-status">
        <span class="status-icon">⚠️</span> 由于错误，使用示例数据进行演示。请检查您的数据文件格式。
    </div>
    """, unsafe_allow_html=True)

# 如果当前使用的是示例数据，显示提示信息
if st.session_state.is_sample_data:
    st.warning("""
    ⚠️ 当前使用的是示例数据，不是您的实际销售数据。要查看真实分析结果：
    1. 请上传您的Excel文件，或
    2. 在左侧边栏设置正确的默认文件路径
    """)

# 显示数据预览
with st.expander("数据预览", expanded=False):
    st.write("以下是加载的数据的前几行：")
    st.dataframe(df.head())
    st.write(f"总行数: {len(df)}")
    st.write(f"列名: {', '.join(df.columns)}")

# 定义新品产品代码
new_products = ['F0110C', 'F0183F', 'F01K8A', 'F0183K', 'F0101P']
new_products_df = df[df['产品代码'].isin(new_products)]

# 创建产品代码到简化名称的映射字典（用于图表显示）
product_name_mapping = {
    code: df[df['产品代码'] == code]['简化产品名称'].iloc[0] if len(df[df['产品代码'] == code]) > 0 else code
    for code in df['产品代码'].unique()
}

# 侧边栏 - 筛选器
st.sidebar.markdown('<div class="sidebar-header">筛选数据</div>', unsafe_allow_html=True)

# 筛选器容器开始
st.sidebar.markdown('<div class="filter-container">', unsafe_allow_html=True)

# 区域筛选器
all_regions = sorted(df['所属区域'].astype(str).unique())
selected_regions = st.sidebar.multiselect("选择区域", all_regions, default=all_regions)

# 客户筛选器
all_customers = sorted(df['客户简称'].astype(str).unique())
selected_customers = st.sidebar.multiselect("选择客户", all_customers, default=[])

# 产品代码筛选器
all_products = sorted(df['产品代码'].astype(str).unique())
product_options = [(code, product_name_mapping[code]) for code in all_products]
selected_products = st.sidebar.multiselect(
    "选择产品",
    options=all_products,
    format_func=lambda x: f"{x} ({product_name_mapping[x]})",
    default=[]
)

# 申请人筛选器
all_applicants = sorted(df['申请人'].astype(str).unique())
selected_applicants = st.sidebar.multiselect("选择申请人", all_applicants, default=[])

# 筛选器容器结束
st.sidebar.markdown('</div>', unsafe_allow_html=True)

# 应用筛选条件
filtered_df = df.copy()

if selected_regions:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['所属区域'].isin(selected_regions)])

if selected_customers:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['客户简称'].isin(selected_customers)])

if selected_products:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['产品代码'].isin(selected_products)])

if selected_applicants:
    filtered_df = safe_filter(filtered_df, lambda d: d[d['申请人'].isin(selected_applicants)])

# 检查筛选后是否还有数据
if filtered_df.empty:
    st.error("应用所有筛选条件后没有匹配的数据。请调整筛选条件。")
    # 重置为原始数据
    filtered_df = df.copy()
    st.warning("已重置为原始数据。")

# 根据筛选后的数据筛选新品数据
filtered_new_products_df = filtered_df[filtered_df['产品代码'].isin(new_products)]

# 导航栏
st.markdown('<div class="sub-header">导航</div>', unsafe_allow_html=True)
tabs = st.tabs(["销售概览", "新品分析", "客户细分", "产品组合", "市场渗透率"])

with tabs[0]:  # 销售概览
    # KPI指标行
    st.markdown('<div class="sub-header">🔑 关键绩效指标</div>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)

    try:
        total_sales = filtered_df['销售额'].sum()
        with col1:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">总销售额</div>
                <div class="metric-value">{format_yuan(total_sales)}</div>
            </div>
            """, unsafe_allow_html=True)

        total_customers = filtered_df['客户简称'].nunique()
        with col2:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">客户数量</div>
                <div class="metric-value">{total_customers}</div>
            </div>
            """, unsafe_allow_html=True)

        total_products = filtered_df['产品代码'].nunique()
        with col3:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">产品数量</div>
                <div class="metric-value">{total_products}</div>
            </div>
            """, unsafe_allow_html=True)

        avg_price = filtered_df['单价（箱）'].mean()
        with col4:
            st.markdown(f"""
            <div class="card">
                <div class="metric-label">平均单价</div>
                <div class="metric-value">{avg_price:.2f}元</div>
            </div>
            """, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"计算KPI指标时出错: {str(e)}")

    # 区域销售分析
    st.markdown('<div class="sub-header section-gap">📊 区域销售分析</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    try:
        # 区域销售额柱状图
        region_sales = filtered_df.groupby('所属区域')['销售额'].sum().reset_index()

        if not region_sales.empty:
            with col1:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                fig_region = px.bar(
                    region_sales,
                    x='所属区域',
                    y='销售额',
                    color='所属区域',
                    title='各区域销售额',
                    labels={'销售额': '销售额 (元)', '所属区域': '区域'},
                    height=500,
                    color_discrete_sequence=px.colors.qualitative.Bold
                )
                # 添加文本标签
                fig_region.update_traces(
                    text=[format_yuan(val) for val in region_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_region.update_layout(
                    xaxis_title=dict(text="区域", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_region.update_yaxes(
                    range=[0, region_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_region, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)

            with col2:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                # 区域销售占比饼图
                fig_region_pie = px.pie(
                    region_sales,
                    values='销售额',
                    names='所属区域',
                    title='各区域销售占比',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Bold
                )
                fig_region_pie.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    textfont=dict(size=14)
                )
                fig_region_pie.update_layout(
                    margin=dict(t=60, b=60, l=60, r=60),
                    font=dict(size=14)
                )
                st.plotly_chart(fig_region_pie, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.warning("没有足够的区域销售数据来创建图表。")
    except Exception as e:
        st.error(f"创建区域销售分析图表时出错: {str(e)}")

    # 产品销售分析
    st.markdown('<div class="sub-header section-gap">📦 产品销售分析</div>', unsafe_allow_html=True)


    # 提取包装类型
    def extract_packaging(product_name):
        try:
            if '袋装' in product_name:
                return '袋装'
            elif '盒装' in product_name:
                return '盒装'
            elif '随手包' in product_name:
                return '随手包'
            elif '迷你包' in product_name:
                return '迷你包'
            elif '分享装' in product_name:
                return '分享装'
            else:
                return '其他'
        except:
            return '其他'


    try:
        filtered_df['包装类型'] = filtered_df['产品名称'].apply(extract_packaging)
        packaging_sales = filtered_df.groupby('包装类型')['销售额'].sum().reset_index()

        col1, col2 = st.columns(2)

        if not packaging_sales.empty:
            with col1:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                # 包装类型销售额柱状图
                fig_packaging = px.bar(
                    packaging_sales.sort_values(by='销售额', ascending=False),
                    x='包装类型',
                    y='销售额',
                    color='包装类型',
                    title='不同包装类型销售额',
                    labels={'销售额': '销售额 (元)', '包装类型': '包装类型'},
                    height=500
                )
                # 添加文本标签
                fig_packaging.update_traces(
                    text=[format_yuan(val) for val in packaging_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_packaging.update_layout(
                    xaxis_title=dict(text="包装类型", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_packaging.update_yaxes(
                    range=[0, packaging_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_packaging, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            # 添加图表容器
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)

            # 价格-销量散点图
            try:
                fig_price_qty = px.scatter(
                    filtered_df,
                    x='单价（箱）',
                    y='数量（箱）',
                    size='销售额',
                    color='所属区域',
                    hover_name='简化产品名称',  # 使用简化产品名称
                    title='价格与销售数量关系',
                    labels={'单价（箱）': '单价 (元/箱)', '数量（箱）': '销售数量 (箱)'},
                    height=500
                )

                # 添加趋势线
                fig_price_qty.update_layout(
                    xaxis_title=dict(text="单价 (元/箱)", font=dict(size=16)),
                    yaxis_title=dict(text="销售数量 (箱)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_price_qty, use_container_width=True)
            except Exception as e:
                st.error(f"创建价格-销量散点图时出错: {str(e)}")

            st.markdown('</div>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"创建产品销售分析图表时出错: {str(e)}")

    # 申请人销售业绩
    st.markdown('<div class="sub-header section-gap">👨‍💼 申请人销售业绩</div>', unsafe_allow_html=True)

    try:
        applicant_performance = filtered_df.groupby('申请人')['销售额'].sum().sort_values(ascending=False).reset_index()

        if not applicant_performance.empty:
            # 添加图表容器
            st.markdown('<div class="chart-container">', unsafe_allow_html=True)

            fig_applicant = px.bar(
                applicant_performance,
                x='申请人',
                y='销售额',
                color='申请人',
                title='申请人销售业绩排名',
                labels={'销售额': '销售额 (元)', '申请人': '申请人'},
                height=500
            )
            # 添加文本标签
            fig_applicant.update_traces(
                text=[format_yuan(val) for val in applicant_performance['销售额']],
                textposition='outside',
                textfont=dict(size=14)
            )
            fig_applicant.update_layout(
                xaxis_title=dict(text="申请人", font=dict(size=16)),
                yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                xaxis_tickfont=dict(size=14),
                yaxis_tickfont=dict(size=14),
                margin=dict(t=60, b=80, l=80, r=60),
                plot_bgcolor='rgba(0,0,0,0)'
            )
            # 确保Y轴有足够空间显示数据标签
            fig_applicant.update_yaxes(
                range=[0, applicant_performance['销售额'].max() * 1.2]
            )
            st.plotly_chart(fig_applicant, use_container_width=True)

            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.warning("没有足够的申请人销售数据来创建图表。")
    except Exception as e:
        st.error(f"创建申请人销售业绩图表时出错: {str(e)}")

    # 原始数据表
    with st.expander("查看筛选后的原始数据"):
        st.dataframe(filtered_df)

with tabs[1]:  # 新品分析
    st.markdown('<div class="sub-header">🆕 新品销售分析</div>', unsafe_allow_html=True)

    # 检查新品数据是否为空
    if filtered_new_products_df.empty:
        st.warning("当前筛选条件下没有新品销售数据。请调整筛选条件或确认产品代码是否正确。")
    else:
        # 新品KPI指标
        col1, col2, col3 = st.columns(3)

        try:
            new_products_sales = filtered_new_products_df['销售额'].sum()
            with col1:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">新品销售额</div>
                    <div class="metric-value">{format_yuan(new_products_sales)}</div>
                </div>
                """, unsafe_allow_html=True)

            new_products_percentage = (new_products_sales / total_sales * 100) if total_sales > 0 else 0
            with col2:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">新品销售占比</div>
                    <div class="metric-value">{new_products_percentage:.2f}%</div>
                </div>
                """, unsafe_allow_html=True)

            new_products_customers = filtered_new_products_df['客户简称'].nunique()
            with col3:
                st.markdown(f"""
                <div class="card">
                    <div class="metric-label">购买新品的客户数</div>
                    <div class="metric-value">{new_products_customers}</div>
                </div>
                """, unsafe_allow_html=True)
        except Exception as e:
            st.error(f"计算新品KPI指标时出错: {str(e)}")

        # 新品销售详情
        st.markdown('<div class="sub-header section-gap">各新品销售额对比</div>', unsafe_allow_html=True)

        try:
            # 使用简化产品名称
            product_sales = filtered_new_products_df.groupby(['产品代码', '简化产品名称'])['销售额'].sum().reset_index()
            product_sales = product_sales.sort_values('销售额', ascending=False)

            if not product_sales.empty:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                fig_product_sales = px.bar(
                    product_sales,
                    x='简化产品名称',  # 使用简化产品名称
                    y='销售额',
                    color='简化产品名称',  # 使用简化产品名称
                    title='新品产品销售额对比',
                    labels={'销售额': '销售额 (元)', '简化产品名称': '产品名称'},
                    height=500
                )
                # 添加文本标签
                fig_product_sales.update_traces(
                    text=[format_yuan(val) for val in product_sales['销售额']],
                    textposition='outside',
                    textfont=dict(size=14)
                )
                fig_product_sales.update_layout(
                    xaxis_title=dict(text="产品名称", font=dict(size=16)),
                    yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                    xaxis_tickfont=dict(size=14),
                    yaxis_tickfont=dict(size=14),
                    margin=dict(t=60, b=80, l=80, r=60),
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                # 确保Y轴有足够空间显示数据标签
                fig_product_sales.update_yaxes(
                    range=[0, product_sales['销售额'].max() * 1.2]
                )
                st.plotly_chart(fig_product_sales, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("没有足够的新品销售数据来创建图表。")
        except Exception as e:
            st.error(f"创建新品销售对比图表时出错: {str(e)}")

        # 区域新品销售分析
        st.markdown('<div class="sub-header section-gap">区域新品销售分析</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)

        try:
            with col1:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                # 区域新品销售额堆叠柱状图
                region_product_sales = filtered_new_products_df.groupby(['所属区域', '简化产品名称'])[
                    '销售额'].sum().reset_index()

                if not region_product_sales.empty:
                    fig_region_product = px.bar(
                        region_product_sales,
                        x='所属区域',
                        y='销售额',
                        color='简化产品名称',  # 使用简化产品名称
                        title='各区域新品销售额分布',
                        labels={'销售额': '销售额 (元)', '所属区域': '区域', '简化产品名称': '产品名称'},
                        height=500
                    )
                    fig_region_product.update_layout(
                        xaxis_title=dict(text="区域", font=dict(size=16)),
                        yaxis_title=dict(text="销售额 (元)", font=dict(size=16)),
                        xaxis_tickfont=dict(size=14),
                        yaxis_tickfont=dict(size=14),
                        margin=dict(t=60, b=80, l=80, r=60),
                        plot_bgcolor='rgba(0,0,0,0)',
                        legend_title="产品名称",
                        legend_font=dict(size=12)
                    )
                    st.plotly_chart(fig_region_product, use_container_width=True)
                else:
                    st.warning("没有足够的区域新品销售数据来创建图表。")

                st.markdown('</div>', unsafe_allow_html=True)

            with col2:
                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                # 新品占比饼图
                fig_new_vs_old = px.pie(
                    values=[new_products_sales, total_sales - new_products_sales],
                    names=['新品', '非新品'],
                    title='新品销售额占总销售额比例',
                    hole=0.4,
                    color_discrete_sequence=['#ff9999', '#66b3ff']
                )
                fig_new_vs_old.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    textfont=dict(size=14)
                )
                fig_new_vs_old.update_layout(
                    margin=dict(t=60, b=60, l=60, r=60),
                    font=dict(size=14)
                )
                st.plotly_chart(fig_new_vs_old, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"创建区域新品销售分析图表时出错: {str(e)}")

        # 区域内新品销售占比热力图
        st.markdown('<div class="sub-header section-gap">各区域内新品销售占比</div>', unsafe_allow_html=True)

        try:
            # 计算各区域的新品总销售额
            region_total_sales = filtered_new_products_df.groupby('所属区域')['销售额'].sum().reset_index()

            # 计算各区域各新品的销售占比
            region_product_sales = filtered_new_products_df.groupby(['所属区域', '产品代码', '简化产品名称'])[
                '销售额'].sum().reset_index()

            if not region_total_sales.empty and not region_product_sales.empty:
                region_product_sales = region_product_sales.merge(region_total_sales, on='所属区域',
                                                                  suffixes=('', '_区域总计'))
                region_product_sales['销售占比'] = region_product_sales['销售额'] / region_product_sales[
                    '销售额_区域总计'] * 100

                # 创建显示名称列（简化产品名称）
                region_product_sales['显示名称'] = region_product_sales['简化产品名称']

                # 透视表
                pivot_percentage = pd.pivot_table(
                    region_product_sales,
                    values='销售占比',
                    index='所属区域',
                    columns='显示名称',  # 使用简化名称作为列名
                    fill_value=0
                )

                # 添加图表容器
                st.markdown('<div class="chart-container">', unsafe_allow_html=True)

                # 使用Plotly创建热力图
                fig_heatmap = px.imshow(
                    pivot_percentage,
                    labels=dict(x="产品名称", y="区域", color="销售占比 (%)"),
                    x=pivot_percentage.columns,
                    y=pivot_percentage.index,
                    color_continuous_scale="YlGnBu",
                    title="各区域内新品销售占比 (%)",
                    height=500
                )

                fig_heatmap.update_layout(
                    xaxis_title=dict(text="产品名称", font=dict(size=16)),
                    yaxis_title=dict(text="区域", font=dict(size=16)),
                    margin=dict(t=80, b=80, l=100, r=100),
                    font=dict(size=14)
                )

                # 添加注释
                for i in range(len(pivot_percentage.index)):
                    for j in range(len(pivot_percentage.columns)):
                        fig_heatmap.add_annotation(
                            x=j,
                            y=i,
                            text=f"{pivot_percentage.iloc[i, j]:.1f}%",
                            showarrow=False,
                            font=dict(color="black" if pivot_percentage.iloc[i, j] < 50 else "white", size=14)
                        )

                st.plotly_chart(fig_heatmap, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("没有足够的区域内新品销售数据来创建热力图。")
        except Exception as e:
            st.error(f"创建区域内新品销售占比热力图时出错: {str(e)}")

        # 新品数据表
        with st.expander("查看新品销售数据"):
            display_columns = [col for col in filtered_new_products_df.columns if
                               col != '产品代码' or col != '产品名称']
            st.dataframe(filtered_new_products_df[display_columns])

# 与原始代码其余部分相同，这里省略其余Tab的代码...
# 客户细分、产品组合和市场渗透率Tab的代码保持不变

# 底部下载区域
st.markdown("---")
st.markdown('<div class="sub-header">📊 导出分析结果</div>', unsafe_allow_html=True)


# 创建Excel报告
def generate_excel_report(df, new_products_df):
    try:
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')

        # 销售概览表
        df.to_excel(writer, sheet_name='销售数据总览', index=False)

        # 新品分析表
        if not new_products_df.empty:
            new_products_df.to_excel(writer, sheet_name='新品销售数据', index=False)

        # 区域销售汇总
        region_summary = df.groupby('所属区域').agg({
            '销售额': 'sum',
            '客户简称': pd.Series.nunique,
            '产品代码': pd.Series.nunique,
            '数量（箱）': 'sum'
        }).reset_index()
        region_summary.columns = ['区域', '销售额', '客户数', '产品数', '销售数量']
        region_summary.to_excel(writer, sheet_name='区域销售汇总', index=False)

        # 产品销售汇总
        product_summary = df.groupby(['产品代码', '简化产品名称']).agg({
            '销售额': 'sum',
            '客户简称': pd.Series.nunique,
            '数量（箱）': 'sum'
        }).sort_values('销售额', ascending=False).reset_index()
        product_summary.columns = ['产品代码', '产品名称', '销售额', '购买客户数', '销售数量']
        product_summary.to_excel(writer, sheet_name='产品销售汇总', index=False)

        # 保存Excel
        writer.close()

        return output.getvalue()
    except Exception as e:
        st.error(f"生成Excel报告时出错: {str(e)}")
        # 返回一个简单的错误报告
        error_output = BytesIO()
        with pd.ExcelWriter(error_output, engine='xlsxwriter') as writer:
            pd.DataFrame({'错误': [f"生成报告时出错: {str(e)}"]}).to_excel(writer, sheet_name='错误信息', index=False)
        return error_output.getvalue()


# 下载按钮
try:
    excel_report = generate_excel_report(filtered_df, filtered_new_products_df)

    st.markdown('<div class="download-button">', unsafe_allow_html=True)
    st.download_button(
        label="下载Excel分析报告",
        data=excel_report,
        file_name="销售数据分析报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown('</div>', unsafe_allow_html=True)
except Exception as e:
    st.error(f"创建下载按钮时出错: {str(e)}")

# 使用说明与分享提示
st.markdown("""
<div class="highlight">
    <h3>📋 使用说明</h3>
    <p><strong>分享仪表盘：</strong> 当您分享此仪表盘链接时，其他人打开链接将会看到基于您配置的默认文件的数据分析。</p>
    <p><strong>自定义分析：</strong> 其他人仍可以上传自己的文件进行分析，但默认会显示您设置的数据。</p>
    <p><strong>更新默认文件：</strong> 您可以在侧边栏的"默认文件设置"中随时更改默认数据源。</p>
</div>
""", unsafe_allow_html=True)

# 底部注释
st.markdown("""
<div style="text-align: center; margin-top: 30px; color: #666;">
    <p>销售数据分析仪表盘 © 2025</p>
</div>
""", unsafe_allow_html=True)
