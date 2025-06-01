import pandas as pd

def get_contact_list(excel_path):
    """获取通信录表格内容"""
    # 使用pandas读取Excel文件
    df = pd.read_excel(excel_path)
    
    # 将DataFrame转换为字典列表
    contact_list = df.to_dict('records')
    
    # 获取实际数据行数(不含表头)
    total_rows = len(df)
    
    return contact_list, total_rows
