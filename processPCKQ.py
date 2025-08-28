import pandas as pd
from datetime import datetime

def process_pc_attendance(file_path):
    """
    处理PC考勤表格数据
    :param file_path: Excel文件路径
    :return: 日期范围(开始日期,结束日期), 精简后的考勤数据DataFrame
    """
    try:
        # 读取Excel文件，可能是csv文件
        df = pd.read_csv(file_path, encoding='gbk')
        # if file_path.endswith('.xlsx'):
        #     df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
        # else:
        #     print("文件不是xlsx格式，尝试读取csv文件")
        #     df = pd.read_csv(file_path)

        # 如果文件中不存在目标列名，给出明确提示
        required_columns = ['姓名', '工号', '出勤状态', '考勤日期', '上班考勤时间', '下班考勤时间']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"缺少必要列：{missing_cols}")

        # 提取所需字段
        df = df[required_columns].copy()

        # 处理考勤日期为datetime格式
        df['考勤日期'] = pd.to_datetime(df['考勤日期'], errors='coerce')

        # 丢弃无效日期
        df = df[df['考勤日期'].notna()]

        if df.empty:
            raise ValueError("未找到有效的考勤日期")

        # 获取起止时间
        start_date = df['考勤日期'].min()
        end_date = df['考勤日期'].max()

        return (start_date, end_date), df

    except Exception as e:
        print(f"处理PC考勤数据时发生错误: {str(e)}")
        return None, None

def get_date_range(file_path):
    """仅获取日期范围"""
    date_range, _ = process_pc_attendance(file_path)
    return date_range

def get_attendance_data(file_path):
    """仅获取考勤数据（精简字段）"""
    _, attendance_data = process_pc_attendance(file_path)
    return attendance_data

# 填充考勤对象的数据
def fill_pc_attendance(index_map, pc_df):
    """
    将PC考勤数据写入模板
    :param index_map: (工号, 日期) -> record 的索引
    :param pc_df: 原始PC考勤DataFrame
    :return: None（直接修改记录）
    """
    for _, row in pc_df.iterrows():
        # 转换为字符串
        emp_id= str(row["工号"]).strip().zfill(8)
        date = pd.to_datetime(row["考勤日期"]).date()
        status = row["出勤状态"]
        start = row["上班考勤时间"]
        end = row["下班考勤时间"]

        key = (emp_id, date)
        # if emp_id == "02005006":
        #     print(f"[DEBUG] 工号{emp_id} 在 {date} 的 start={repr(start)} ({type(start)}), end={repr(end)} ({type(end)})")

        if is_empty_time(start) and is_empty_time(end):
            if key in index_map:
                index_map[key]["pc出勤状态"] = ""
        elif key in index_map:
            index_map[key]["pc出勤状态"] = status

def is_empty_time(val):
    if val is None:
        return True
    if isinstance(val, str):
        return val.strip() == ''
    return pd.isna(val)
