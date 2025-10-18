import streamlit as st
import pandas as pd
from collections import defaultdict
import re
import uuid
import os

# 设置页面配置
st.set_page_config(page_title="SP-批量模版生成工具 - JP版", page_icon="📊", layout="centered")

# 自定义 CSS 样式
st.markdown("""
    <style>
    /* 主标题样式 */
    .main-title {
        font-size: 2.5em;
        font-weight: bold;
        color: #2C3E50;
        text-align: center;
        margin-bottom: 20px;
    }
    /* 提示文字样式 */
    .instruction {
        font-size: 1.1em;
        color: #34495E;
        margin-bottom: 20px;
    }
    /* 按钮样式 */
    .stButton>button {
        background-color: #3498DB;
        color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 1em;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2980B9;
    }
    /* 文件上传框样式 */
    .stFileUploader label {
        font-size: 1.1em;
        color: #2C3E50;
        font-weight: bold;
    }
    /* 成功和错误消息样式 */
    .stSuccess {
        background-color: #E8F5E9;
        border-left: 5px solid #4CAF50;
        padding: 10px;
        border-radius: 5px;
    }
    .stError {
        background-color: #FFEBEE;
        border-left: 5px solid #F44336;
        padding: 10px;
        border-radius: 5px;
    }
    .stWarning {
        background-color: #FFF3E0;
        border-left: 5px solid #FF9800;
        padding: 10px;
        border-radius: 5px;
    }
    </style>
""", unsafe_allow_html=True)

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
    
    # 改进的关键词类别提取逻辑
    def extract_keyword_categories(df_survey):
        categories = set()
        
        # 从列名中提取所有可能的关键词类别
        for col in df_survey.columns:
            col_lower = str(col).lower()
            
            # 处理关键词列
            if any(x in col_lower for x in ['精准词', '广泛词', '精准', '广泛']):
                # 去除匹配类型后缀
                for suffix in ['精准词', '广泛词', '精准', '广泛']:
                    if col_lower.endswith(suffix):
                        prefix = col_lower[:-len(suffix)].strip()
                        # 按多种分隔符拆分
                        parts = re.split(r'[/\-_\s\.]', prefix)
                        for part in parts:
                            if part and len(part) > 1:  # 只保留有意义的词
                                categories.add(part)
                        break
            
            # 处理ASIN列
            elif 'asin' in col_lower and '否定' not in col_lower:
                # 去除ASIN后缀
                prefix = col_lower.replace('asin', '').strip()
                # 按多种分隔符拆分
                parts = re.split(r'[/\-_\s\.]', prefix)
                for part in parts:
                    if part and len(part) > 1:  # 只保留有意义的词
                        categories.add(part)
        
        # 添加已知的关键词类别
        categories.update(['suzhu', '宿主', 'case', '包', 'tape'])
        
        # 移除空字符串
        categories.discard('')
        
        return categories
    
    keyword_categories = extract_keyword_categories(df_survey)
    st.write(f"识别到的关键词类别: {keyword_categories}")
    
    # 生成数据行
    rows = []
    
    # 函数：为商品定向活动查找匹配的列
    def find_matching_asin_columns(campaign_name, df_survey, keyword_categories):
        campaign_name_normalized = str(campaign_name).lower()
        
        # 1. 检查是否为ASIN活动
        if 'asin' not in campaign_name_normalized:
            st.write(f"  {campaign_name} 不是商品定向活动，无匹配ASIN列")
            return []
        
        # 2. 确定关键词类别
        matched_categories = []
        for category in keyword_categories:
            if category and category in campaign_name_normalized:
                matched_categories.append(category)
        
        if not matched_categories:
            st.write(f"  {campaign_name} 未匹配到任何关键词类别，无匹配ASIN列")
            return []
        
        st.write(f"  匹配的关键词类别: {matched_categories}")
        
        # 3. 查找匹配的ASIN列
        matching_columns = []
        for col in df_survey.columns:
            col_lower = str(col).lower()
            if 'asin' in col_lower and '否定' not in col_lower:
                # 检查列名是否包含任何匹配的类别
                for category in matched_categories:
                    if category in col_lower:
                        matching_columns.append(col)
                        break
        
        st.write(f"  初步匹配的ASIN列: {matching_columns}")
        
        # 4. 如果有多个匹配列，进一步根据活动名称和列名中共同的词筛选
        if len(matching_columns) > 1:
            # 提取活动名称中的所有词
            campaign_words = campaign_name_normalized.split()
            st.write(f"  调试: 提取的campaign_words: {campaign_words}")  # Debug print to verify
            
            # 排除已经用于匹配的类别词和常见停用词
            exclude_words = set(matched_categories + ['asin', '商品定向', '定向', '精准', '广泛', 'exact', 'broad', 'sp'])
            candidate_words = [word for word in campaign_words if word and word not in exclude_words]
            
            st.write(f"  候选匹配词: {candidate_words}")
            
            # 查找包含候选词的列
            word_matched_columns = []
            for col in matching_columns:
                col_lower = str(col).lower()
                # 检查列名是否包含任何候选词
                for word in candidate_words:
                    if word in col_lower:
                        word_matched_columns.append((col, word))
                        st.write(f"    列 '{col}' 包含共同词 '{word}'")
                        break
            
            if word_matched_columns:
                # 如果有多个列匹配到候选词，选择第一个
                if len(word_matched_columns) > 1:
                    st.write(f"  多个列匹配到候选词，选择第一个: {word_matched_columns[0][0]}")
                matching_columns = [word_matched_columns[0][0]]
            else:
                # 如果没有列匹配到候选词，选择第一个匹配的列
                st.write(f"  没有列匹配到候选词，选择第一个匹配列: {matching_columns[0]}")
                matching_columns = [matching_columns[0]]
        
        return matching_columns

    # 函数：查找匹配的关键词列
    def find_matching_keyword_columns(campaign_name, df_survey, keyword_categories, keyword_columns, match_type):
        campaign_name_normalized = str(campaign_name).lower()
        
        # 确定关键词类别
        matched_categories = []
        for category in keyword_categories:
            if category and category in campaign_name_normalized:
                matched_categories.append(category)
        
        st.write(f"  匹配的关键词类别: {matched_categories}")
        
        if not matched_categories:
            st.write("  无匹配的关键词类别")
            return [], []
        
        # 确定匹配类型关键词
        match_type_keywords = []
        if match_type == '精准':
            match_type_keywords = ['精准', 'exact']
        elif match_type == '广泛':
            match_type_keywords = ['广泛', 'broad']
        
        # 查找匹配的列
        matching_columns = []
        for col in keyword_columns:
            col_lower = str(col).lower()
            
            # 检查列名是否包含匹配类型关键词
            has_match_type = any(keyword in col_lower for keyword in match_type_keywords)
            
            # 检查列名是否包含任何匹配的类别
            has_category = any(category in col_lower for category in matched_categories)
            
            if has_match_type and has_category:
                matching_columns.append(col)
        
        st.write(f"  匹配的列: {matching_columns}")
        
        # 提取关键词
        keywords = []
        for col in matching_columns:
            keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        
        # 去重
        keywords = list(dict.fromkeys(keywords))
        st.write(f"  关键词数量: {len(keywords)} (示例: {keywords[:2] if keywords else '无'})")
        
        return matching_columns, keywords
    
    # 函数：查找否定关键词
    def find_neg_keywords(campaign_name, df_survey, keyword_categories, keyword_columns):
        campaign_name_normalized = str(campaign_name).lower()
        
        # 确定关键词类别
        matched_categories = []
        for category in keyword_categories:
            if category and category in campaign_name_normalized:
                matched_categories.append(category)
        
        if not matched_categories:
            return []
        
        # 查找精准关键词列
        neg_keywords = []
        for col in keyword_columns:
            col_lower = str(col).lower()
            if any(category in col_lower for category in matched_categories) and any(x in col_lower for x in ['精准', 'exact']):
                neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        
        # 去重
        neg_keywords = list(dict.fromkeys(neg_keywords))
        st.write(f"  精准否定关键词数量: {len(neg_keywords)} (示例: {neg_keywords[:2] if neg_keywords else '无'})")
        
        return neg_keywords
    
    # 函数：查找交叉否定关键词
    def find_cross_neg_keywords(campaign_name, df_survey, keyword_categories, keyword_columns):
        campaign_name_normalized = str(campaign_name).lower()
        
        cross_neg_keywords = []
        
        # 如果是宿主组，否定case精准词
        if any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
            # 查找case精准关键词作为否定词
            for col in keyword_columns:
                col_lower = str(col).lower()
                if any(case_word in col_lower for case_word in ['case', '包', 'tape']) and any(x in col_lower for x in ['精准', 'exact']):
                    cross_neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        
        # 如果是case组，否定宿主精准词
        elif any(x in campaign_name_normalized for x in ['case', '包', 'tape']):
            # 查找宿主精准关键词作为否定词
            for col in keyword_columns:
                col_lower = str(col).lower()
                if any(suzhu_word in col_lower for suzhu_word in ['suzhu', '宿主']) and any(x in col_lower for x in ['精准', 'exact']):
                    cross_neg_keywords.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
        
        # 去重
        cross_neg_keywords = list(dict.fromkeys(cross_neg_keywords))
        st.write(f"  交叉否定关键词数量: {len(cross_neg_keywords)} (示例: {cross_neg_keywords[:2] if cross_neg_keywords else '无'})")
        
        return cross_neg_keywords
    
    for campaign_name in unique_campaigns:
        # 获取 CPC、SKU、广告组默认竞价、预算
        if campaign_name in campaign_to_values:
            cpc = campaign_to_values[campaign_name]['CPC']
            sku = campaign_to_values[campaign_name]['SKU']
            group_bid = campaign_to_values[campaign_name]['广告组默认竞价']
            budget = campaign_to_values[campaign_name]['预算']
        else:
            cpc = 0.5
            sku = 'SKU-1'
            group_bid = default_group_bid
            budget = default_daily_budget
        
        st.write(f"处理活动: {campaign_name}")
        
        campaign_name_normalized = str(campaign_name).lower()
        
        # 确定匹配类型
        is_exact = any(x in campaign_name_normalized for x in ['精准', 'exact'])
        is_broad = any(x in campaign_name_normalized for x in ['广泛', 'broad'])
        is_asin = 'asin' in campaign_name_normalized
        match_type = '精准' if is_exact else '广泛' if is_broad else 'ASIN' if is_asin else None
        st.write(f"  is_exact: {is_exact}, is_broad: {is_broad}, is_asin: {is_asin}, match_type: {match_type}")
        
        # 提取关键词（用于正向关键词，精准/广泛匹配）
        keywords = []
        matched_columns = []
        if is_exact or is_broad:
            matched_columns, keywords = find_matching_keyword_columns(
                campaign_name, df_survey, keyword_categories, keyword_columns, match_type
            )
        
        # 提取精准关键词（用于广泛匹配活动的否定关键词）
        neg_keywords = []
        if is_broad:
            neg_keywords = find_neg_keywords(campaign_name, df_survey, keyword_categories, keyword_columns)
        
        # 提取 ASIN（用于商品定向）
        asin_targets = []
        if is_asin:
            matching_columns = find_matching_asin_columns(campaign_name, df_survey, keyword_categories)
            for col in matching_columns:
                asin_targets.extend([kw for kw in df_survey[col].dropna() if str(kw).strip()])
            asin_targets = list(dict.fromkeys(asin_targets))
            st.write(f"  商品定向 ASIN 数量: {len(asin_targets)} (示例: {asin_targets[:2] if asin_targets else '无'})")
        
        # 广告活动行
        rows.append([
            product, '广告活动', operation, campaign_name, '', '', '', '', '',
            campaign_name, '', '', '', targeting_type, status, budget, '', '',
            '', '', '', bidding_strategy, '', '', ''
        ])
        
        # 广告组行
        rows.append([
            product, '广告组', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', '', group_bid,
            '', '', '', '', '', '', ''
        ])
        
        # 商品广告行
        rows.append([
            product, '商品广告', operation, campaign_name, campaign_name, '', '', '', '',
            campaign_name, campaign_name, '', '', '', status, '', sku, '',
            '', '', '', '', '', '', ''
        ])
        
        # 关键词行（仅精准/广泛匹配）
        if is_exact or is_broad:
            for kw in keywords:
                rows.append([
                    product, '关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, kw, match_type, '', '', '', ''
                ])
        
        # 否定关键词行（仅精准/广泛匹配）
        if is_exact:
            # 原有规则：全局否定精准和否定词组
            for kw in neg_exact:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            for kw in neg_phrase:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定词组', '', '', '', ''
                ])
            
            # 新增：交叉否定规则（宿主精准组否定case精准词，case精准组否定宿主精准词）
            cross_neg_keywords = find_cross_neg_keywords(campaign_name, df_survey, keyword_categories, keyword_columns)
            for kw in cross_neg_keywords:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
        elif is_broad:
            # 全局否定精准
            for kw in neg_exact:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            # 全局否定词组
            for kw in neg_phrase:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定词组', '', '', '', ''
                ])
            # 同类的精准关键词作为否定精准
            for kw in neg_keywords:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
            
            # 交叉否定规则：宿主广泛组否定case精准词，case广泛组否定宿主精准词
            cross_neg_keywords = find_cross_neg_keywords(campaign_name, df_survey, keyword_categories, keyword_columns)
            
            # 如果是宿主广泛组，添加宿主额外否定词
            if any(x in campaign_name_normalized for x in ['suzhu', '宿主']):
                for kw in suzhu_extra_neg_exact:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定精准匹配', '', '', '', ''
                    ])
                for kw in suzhu_extra_neg_phrase:
                    rows.append([
                        product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                        campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                        kw, '否定词组', '', '', '', ''
                    ])
            
            # 添加交叉否定关键词
            for kw in cross_neg_keywords:
                rows.append([
                    product, '否定关键词', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '', '',
                    kw, '否定精准匹配', '', '', '', ''
                ])
        
        # 商品定向和否定商品定向（仅 ASIN 组）
        if is_asin:
            for asin in asin_targets:
                rows.append([
                    product, '商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    cpc, '', '', '', '', '', f'asin="{asin}"'
                ])
            for asin in neg_asin:
                rows.append([
                    product, '否定商品定向', operation, campaign_name, campaign_name, '', '', '', '',
                    campaign_name, campaign_name, '', '', '', status, '', '', '',
                    '', '', '', '', '', '', f'asin="{asin}"'
                ])
    
    # 创建 DataFrame
    df_header = pd.DataFrame(rows, columns=columns)
    try:
        df_header.to_excel(output_file, index=False, engine='openpyxl')
        st.success(f"生成完成！输出文件：{output_file}，总行数：{len(rows)}")
        return output_file
    except PermissionError:
        st.error(f"错误：无法写入 {output_file}，请确保文件未被占用或有写入权限。")
        return None
    except Exception as e:
        st.error(f"写入文件时出错：{e}")
        return None
    
    # 调试输出
    keyword_rows = [row for row in rows if row[1] == '关键词']
    st.write(f"关键词行数量: {len(keyword_rows)}")
    if keyword_rows:
        st.write(f"示例关键词行: 实体层级={keyword_rows[0][1]}, 关键词文本={keyword_rows[0][19]}, 匹配类型={keyword_rows[0][20]}")
    
    product_targeting_rows = [row for row in rows if row[1] == '商品定向']
    st.write(f"商品定向行数量: {len(product_targeting_rows)}")
    if product_targeting_rows:
        st.write(f"示例商品定向行: 实体层级={product_targeting_rows[0][1]}, 竞价={product_targeting_rows[0][18]}, 拓展商品投放编号={product_targeting_rows[0][24]}")
    
    levels = set(row[1] for row in rows)
    st.write(f"所有实体层级: {levels}")

# Streamlit 界面
st.markdown('<div class="main-title">SP-批量模版生成工具 - JP版</div>', unsafe_allow_html=True)
st.markdown('<div class="instruction">上传 survey-JP.xlsx 文件，点击按钮生成 header-JP.xlsx。<br>Upload survey-JP.xlsx file, click button to generate header-JP.xlsx.</div>', unsafe_allow_html=True)

# 文件上传
uploaded_file = st.file_uploader("上传 survey-JP.xlsx", type=["xlsx"])

if uploaded_file is not None:
    # 保存上传的文件
    with open("survey-JP.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # 运行按钮
    if st.button("生成 Header 文件"):
        output_file = 'header-JP.xlsx'
        with st.spinner("正在处理文件..."):
            result = generate_header_from_survey("survey-JP.xlsx", output_file)
            if result and os.path.exists(result):
                with open(result, "rb") as f:
                    st.download_button(
                        label="下载 header-JP.xlsx",
                        data=f,
                        file_name="header-JP.xlsx",
                        mime="application/vnd.openpyxlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("生成文件失败，请检查上传的文件格式或内容。")
</xaiArtifact>

### 部署说明
1. **保存代码**：
   - 将上述代码保存为 `app.py`（Streamlit 默认入口文件）。
2. **requirements.txt**：
   - 创建 `requirements.txt` 文件，内容：