import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import uuid
import os
from typing import Set

# 改进的关键词类别提取逻辑
def extract_keyword_categories(df_survey):
    """
    从调查 DataFrame 的列名中提取关键词类别。
    
    Args:
        df_survey (pd.DataFrame): 调查数据 DataFrame。
    
    Returns:
        Set[str]: 提取的关键词类别集合。
    """
    categories: Set[str] = set()
    
    # 从列名中提取所有可能的关键词类别
    for col in df_survey.columns:
        col_lower = str(col).lower()
        
        # 处理关键词列（包含精准词、广泛词等）
        if any(x in col_lower for x in ['精准词', '广泛词', '精准', '广泛']):
            # 去除匹配类型后缀
            for suffix in ['精准词', '广泛词', '精准', '广泛']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    # 按多种分隔符拆分（/、-、_、空格、.）
                    parts = re.split(r'[/\-_\s\.]', prefix)
                    for part in parts:
                        if part and len(part) > 1:  # 只保留有意义的词（长度 > 1）
                            categories.add(part)
                    break
        
        # 处理 ASIN 列（排除否定 ASIN）
        elif 'asin' in col_lower and '否定' not in col_lower:
            # 去除 ASIN 后缀
            for suffix in ['asin']:
                if col_lower.endswith(suffix):
                    prefix = col_lower[:-len(suffix)].strip()
                    # 按多种分隔符拆分
                    parts = re.split(r'[/\-_\s\.]', prefix)
                    for part in parts:
                        if part and len(part) > 1:  # 只保留有意义的词
                            categories.add(part)
                    break
    
    return categories

# script-JP.py 的函数
def generate_header_from_survey(survey_file='survey-JP.xlsx', output_file='header-JP.xlsx', sheet_name=0):
    try:
        # 读取 Excel 文件
        df_survey = pd.read_excel(survey_file, sheet_name=sheet_name)
        st.write(f"成功读取文件：{survey_file}，数据形状：{df_survey.shape}")
        st.write(f"列名列表: {list(df_survey.columns)}")
    except FileNotFoundError:
        st.error(f"错误：未找到文件 {survey_file}。请确保文件已上传。")
        return None
    except Exception as e:
        st.error(f"读取文件时出错：{e}")
        return None
    
    # 提取独特活动名称
    unique_campaigns = [name for name in df_survey['广告活动名称'].dropna() if str(name).strip()]
    st.write(f"独特活动名称数量: {len(unique_campaigns)}: {unique_campaigns}")
    
    # 创建活动到 CPC/SKU/广告组默认竞价/预算 的映射
    non_empty_campaigns = df_survey[
        df_survey['广告活动名称'].notna() & 
        (df_survey['广告活动名称'] != '')
    ]
    required_cols = ['CPC', 'SKU', '广告组默认竞价', '预算']
    if all(col in non_empty_campaigns.columns for col in required_cols):
        campaign_to_values = non_empty_campaigns.drop_duplicates(
            subset='广告活动名称', keep='first'
        ).set_index('广告活动名称')[required_cols].to_dict('index')
    else:
        campaign_to_values = {}
        st.warning(f"警告：缺少列 {set(required_cols) - set(non_empty_campaigns.columns)}，使用默认值")
    
    st.write(f"生成的字典（有 {len(campaign_to_values)} 个活动）: {campaign_to_values}")
    
    # 关键词列：第 H 列（索引 7）到第 Q 列（索引 16）
    keyword_columns = df_survey.columns[7:17]
    st.write(f"关键词列: {list(keyword_columns)}")
    
    # 检查关键词重复
    duplicates_found = False
    st.write("### 检查关键词重复")
    for col in keyword_columns:
        col_index = list(df_survey.columns).index(col) + 1
        col_letter = chr(64 + col_index) if col_index <= 26 else f"{chr(64 + (col_index-1)//26)}{chr(64 + (col_index-1)%26 + 1)}"
        kw_list = [kw for kw in df_survey[col].dropna() if str(kw).strip()]
        
        if len(kw_list) > len(set(kw_list)):
            duplicates_mask = df_survey[col].duplicated(keep=False)
            duplicates_df = df_survey[duplicates_mask][[col]].dropna()
            st.warning(f"警告：{col_letter} 列 ({col}) 有重复关键词")
            for _, row in duplicates_df.iterrows():
                kw = str(row[col]).strip()
                count = (df_survey[col] == kw).sum()
                if count > 1:
                    st.write(f"  重复词: '{kw}' (出现 {count} 次)")
            duplicates_found = True
    
    if duplicates_found:
        st.error("提示：由于检测到关键词重复，本次不生成表格。请清理重复后重试。")
        return None
    
    st.write("关键词无重复，继续生成...")
    
    # 否定关键词聚合
    neg_exact = [kw for kw in df_survey.get('否定精准', pd.Series()).dropna() if str(kw).strip()]
    neg_phrase = [kw for kw in df_survey.get('否定词组', pd.Series()).dropna() if str(kw).strip()]
    suzhu_extra_neg_exact = [kw for kw in df_survey.get('宿主额外否精准', pd.Series()).dropna() if str(kw).strip()]
    suzhu_extra_neg_phrase = [kw for kw in df_survey.get('宿主额外否词组', pd.Series()).dropna() if str(kw).strip()]
    neg_asin = [kw for kw in df_survey.get('否定ASIN', pd.Series()).dropna() if str(kw).strip()]
    
    # 列定义
    columns = [
        '产品', '实体层级', '操作', '广告活动编号', '广告组编号', '广告组合编号', '广告编号', '关键词编号', '商品投放 ID',
        '广告活动名称', '广告组名称', '开始日期', '结束日期', '投放类型', '状态', '每日预算', 'SKU', '广告组默认竞价',
        '竞价', '关键词文本', '匹配类型', '竞价方案', '广告位', '百分比', '拓展商品投放编号'
    ]
    
    # 默认值
    product = '商品推广'
    operation = 'Create'
    status = '已启用'
    targeting_type = '手动'
    bidding_strategy = '动态竞价 - 仅降低'
    default_daily_budget = 12
    default_group_bid = 0.6
    
    # 提取关键词类别（使用修复后的函数）
    keyword_categories = extract_keyword_categories(df_survey)
    st.write(f"提取的关键词类别: {keyword_categories}")
    
    # 生成 DataFrame（这里添加示例逻辑，基于你的需求扩展）
    data_list = []
    # 示例：为每个独特活动生成行（实际逻辑需根据需求填充）
    for campaign in unique_campaigns:
        # 获取活动特定值或默认值
        campaign_values = campaign_to_values.get(campaign, {})
        cpc = campaign_values.get('CPC', default_daily_budget)
        sku = campaign_values.get('SKU', '')
        group_bid = campaign_values.get('广告组默认竞价', default_group_bid)
        budget = campaign_values.get('预算', default_daily_budget)
        
        # 示例行（扩展为实际关键词处理）
        row = {
            '产品': product,
            '实体层级': '关键词',
            '操作': operation,
            '广告活动编号': str(uuid.uuid4())[:8],  # 示例 UUID
            '广告组编号': str(uuid.uuid4())[:8],
            '广告组合编号': '',
            '广告编号': '',
            '关键词编号': str(uuid.uuid4())[:8],
            '商品投放 ID': '',
            '广告活动名称': campaign,
            '广告组名称': f"{campaign}_组",  # 示例
            '开始日期': '2025-10-17',  # 当前日期示例
            '结束日期': '2025-12-31',
            '投放类型': targeting_type,
            '状态': status,
            '每日预算': budget,
            'SKU': sku,
            '广告组默认竞价': group_bid,
            '竞价': cpc,
            '关键词文本': '',  # 从关键词列填充
            '匹配类型': 'exact',  # 示例
            '竞价方案': bidding_strategy,
            '广告位': '',
            '百分比': 0,
            '拓展商品投放编号': ''
        }
        data_list.append(row)
    
    # 创建 DataFrame
    df_output = pd.DataFrame(data_list, columns=columns)
    
    # 保存到 Excel
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_output.to_excel(writer, sheet_name='Header', index=False)
        st.success(f"成功生成文件：{output_file}")
        st.write("生成的表格预览：")
        st.dataframe(df_output.head())
        return df_output
    except Exception as e:
        st.error(f"保存文件时出错：{e}")
        return None

# Streamlit 主界面
def main():
    st.title("广告活动表头生成工具 (JP)")
    
    # 文件上传
    uploaded_file = st.file_uploader("上传调查 Excel 文件 (survey-JP.xlsx)", type=['xlsx'])
    
    if uploaded_file is not None:
        # 保存上传的文件
        with open('survey-JP.xlsx', 'wb') as f:
            f.write(uploaded_file.getbuffer())
        st.success("文件上传成功！")
        
        # 生成表头
        if st.button("生成表头"):
            generate_header_from_survey(survey_file='survey-JP.xlsx', output_file='header-JP.xlsx')
    else:
        st.info("请上传 survey-JP.xlsx 文件以开始生成。")

if __name__ == "__main__":
    main()