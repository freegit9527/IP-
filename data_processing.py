import pandas as pd
from ipaddress import ip_address, ip_network

def is_ip_in_range(ip, ip_range):
    """判断单个IP是否位于给定的IP段内"""
    try:
        return ip_address(ip) in ip_network(ip_range)
    except ValueError:
        return False

def is_string_match(str1, str2):
    """判断两个字符串是否相等（大小写敏感）"""
    return str1 == str2

def match_data(file1_path, col1_type, col1, file2_path, col2_type, col2):
    """匹配两个表格中的数据信息并合并数据，支持IP和字符串比较，假设使用第一个sheet"""
    df1 = pd.read_excel(file1_path, sheet_name=0)
    df2 = pd.read_excel(file2_path, sheet_name=0)
    
    # 根据类型选择匹配函数
    match_func = is_ip_in_range if col1_type == 'ip' else is_string_match
    
    # 初始化一个列表用于存储匹配的结果行
    matched_rows = []
    # 初始化一个集合用于存储已经匹配过的索引，避免重复添加
    matched_indices = set()
    
    for index1, row1 in df1.iterrows():
        value_to_match = row1[col1]
        # 查找匹配的值
        matches = df2[df2.iloc[:, col2].apply(lambda x: match_func(value_to_match, str(x)))]
        if not matches.empty:
            # 对于每一个匹配，将两行数据合并到一个字典中
            for index2, row2 in matches.iterrows():
                combined_row = {**row1, **row2}  # 合并两个字典
                matched_rows.append(combined_row)
                # 记录匹配的索引，避免之后重复添加
                matched_indices.add(index1)
    
    # 从df1中筛选出未匹配的行，排除已匹配的索引
    unmatched_rows = df1[~df1.index.isin(matched_indices)].reset_index(drop=True)
    
    # 结果合并，先匹配的，后未匹配的
    result_df = pd.DataFrame(matched_rows)
    result_df = pd.concat([result_df, unmatched_rows], ignore_index=True)
    return result_df