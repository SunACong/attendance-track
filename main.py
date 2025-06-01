import pandas as pd
import os
from datetime import timedelta
from tabulate import tabulate
from processPCKQ import process_pc_attendance, get_date_range, get_attendance_data
from processTXL import get_contact_list


def read_excel_file(file_path, sheet_name=0, usecols=None):
    """
    读取 Excel 文件内容并返回为 JSON 格式数据（字典列表）

    :param file_path: Excel 文件路径
    :param sheet_name: 默认读取第一个工作表，也可指定名称
    :param usecols: 指定要读取的列，可以是列名列表或列索引列表
    :return: List[Dict]，每行为一个 dict
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    try:
        df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            engine='openpyxl',
            usecols=usecols
        )
        # 去除空行
        df = df.dropna(how='all')
        return df.to_dict(orient='records')
    except Exception as e:
        raise Exception(f"读取 Excel 文件失败: {str(e)}")

def init_attendance_template(df, start_date, end_date):
    """
    初始化考勤模板列表（每人每天一条记录）
    :param df: 含姓名、工号、所在部门的DataFrame
    :param start_date: 起始日期
    :param end_date: 结束日期
    :return: 模板列表
    """
    if isinstance(df, pd.DataFrame):
        # 强制工号为字符串类型
        df["工号"] = df["工号"].astype(str).str.zfill(8)
        unique_people = df.drop_duplicates(subset=['姓名', '工号'])
    else:
        # 如果是字典列表，转换为 DataFrame，再强制工号为字符串
        df = pd.DataFrame(df)
        df["工号"] = df["工号"].astype(str).str.zfill(8)
        unique_people = df.drop_duplicates(subset=['姓名', '工号'])

    date_range = pd.date_range(start=start_date, end=end_date).date

    template_records = []
    for _, person in unique_people.iterrows():
        for date in date_range:
            template_records.append({
                "姓名": person["姓名"],
                "工号": person["工号"],
                "部门": person.get("所在部门", ""),  # 使用 .get 更健壮
                "考勤日期": date,
                "pc出勤状态": "",
                "oa出勤状态": "",
                "oa离岗登记": "",
                "oa请假信息": ""
            })
    return template_records


def build_record_index(template_records):
    """
    构建一个 (工号, 日期) -> record 的快速索引
    :param template_records: 模板记录列表
    :return: 索引字典
    """
    return {
        (str(record["工号"]).strip(), record["考勤日期"]): record
        for record in template_records
    }


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

        key = (emp_id, date)
        if key in index_map:
            index_map[key]["pc出勤状态"] = status

def fill_oa_attendance(index_map, oa_df):
    """
    根据 OA 打卡数据填充 oa出勤状态 和 是否打卡
    :param index_map: (工号, 日期) -> record 的索引
    :param oa_df: 原始OA打卡记录（DataFrame）
    """
    # 转换时间字段
    oa_df["打卡时间"] = pd.to_datetime(oa_df["打卡时间"])

    # 添加新列：日期、小时
    oa_df["打卡日期"] = oa_df["打卡时间"].dt.date
    oa_df["打卡小时"] = oa_df["打卡时间"].dt.hour
    oa_df["打卡分钟"] = oa_df["打卡时间"].dt.minute

    # 分组处理：按工号 + 打卡日期聚合 工号转换为字符串
    oa_df["工号"] = oa_df["工号"].astype(str).str.zfill(8)
    grouped = oa_df.groupby(["工号", "打卡日期"])

    for (emp_id, date), group in grouped:
        has_morning = any(t < 8 for t in group["打卡小时"])
        has_evening = any(t > 18 or (t == 18 and m > 0) for t, m in zip(group["打卡小时"], group["打卡分钟"]))

        key = (emp_id, date)
        if key in index_map:
            record = index_map[key]
            record["oa出勤状态"] = "正常出勤" if has_morning and has_evening else "异常"
            record["oa是否打卡"] = True

def fill_leave_registration(index_map, leave_df):
    leave_df.columns = leave_df.columns.str.strip()
    leave_df["离岗日期"] = pd.to_datetime(leave_df["离岗日期"])
    leave_df["返岗日期"] = pd.to_datetime(leave_df["返岗日期"])
    leave_df["工号"] = leave_df["工号"].astype(str).str.strip()

    for _, row in leave_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        start_date = row["离岗日期"].date()
        end_date = row["返岗日期"].date()

        current_date = start_date
        while current_date <= end_date:
            key = (emp_id, current_date)

            if key in index_map:
                index_map[key]["oa离岗登记"] = True
            else:
                print(f"❗未找到 key: {key}，请确认 index_map 中是否存在")
            current_date += timedelta(days=1)

def fill_leave_info(index_map, leave_df):
    """
    根据请假数据更新 index_map 中的 oa请假信息（为 True）
    :param index_map: (工号, 日期) -> record
    :param leave_df: 请假 DataFrame
    """
    # 统一解析日期字段（支持 5/23/25 这种格式）
    leave_df["请假开始日期"] = pd.to_datetime(leave_df["请假开始日期"], errors="coerce")
    leave_df["请假结束日期"] = pd.to_datetime(leave_df["请假结束日期"], errors="coerce")

    for _, row in leave_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        start_date = row["请假开始日期"].date()
        end_date = row["请假结束日期"].date()

        current_date = start_date
        while current_date <= end_date:
            key = (emp_id, current_date)
            if key in index_map:
                record = index_map[key]
                record["oa请假信息"] = True  # ✅ 标记为 True
            current_date += timedelta(days=1)


def summarize_attendance(contact_attendance_list, holiday_set):
    summary_map = {}

    for record in contact_attendance_list:
        print(holiday_set)

        # ✅ 如果是节假日，跳过这一天
        if record["考勤日期"] in holiday_set:
            continue

        name = record.get("姓名")
        emp_id = str(record.get("工号")).strip().zfill(8)
        dept = record.get("部门")

        pc_status = record.get("pc出勤状态")
        oa_status = record.get("oa出勤状态")
        oa_leave = record.get("oa请假信息")
        oa_absence = record.get("oa离岗登记")
        oa_clock = record.get("oa是否打卡")

        # 初始化统计结构
        if emp_id not in summary_map:
            summary_map[emp_id] = {
                "姓名": name,
                "工号": emp_id,
                "部门": dept,
                "正常出勤天数": 0,
                "缺勤天数": 0,
                "请假天数": 0,
                "旷工天数": 0,
            }

        stat = summary_map[emp_id]

        is_all_empty = not pc_status and not oa_status and not oa_absence and not oa_leave and not oa_clock
        is_pc_normal = oa_absence is True or pc_status == "正常出勤"
        is_oa_normal = oa_status == "正常出勤"
        has_oa_leave = oa_leave is True

        if is_all_empty:
            stat["旷工天数"] += 1
        elif has_oa_leave:
            stat["请假天数"] += 1
        elif is_pc_normal or is_oa_normal:
            stat["正常出勤天数"] += 1
        else:
            stat["缺勤天数"] += 1

    # 返回聚合后的列表（可用于生成表格/导出）
    return list(summary_map.values())

if __name__ == "__main__":
    holiday_set = set()
    try:
        person_data = read_excel_file("cs/01_通信录.xlsx", usecols=['姓名', '工号','所在部门'])
        oa_data = pd.read_excel("cs/02_移动OA_考勤_cs.xlsx", engine="openpyxl")
        leave_df = pd.read_excel("cs/03_移动OA_离岗登记表_cs.xlsx", engine="openpyxl")
        qj_df = pd.read_excel("cs/04_移动OA_请假信息统计_cs.xlsx", engine="openpyxl")
        holiday_df = pd.read_excel("cs/05_节假日表.xlsx")
        holiday_set = set(pd.to_datetime(holiday_df["日期"]).dt.date)
        print(holiday_set)
    except Exception as e:
        print(e)
    
    # 调用processPCKQ.py中的函数
    pc_attendance_path = "cs/00_PC_考勤结果_cs.xlsx"
    txl_attendance_path = "cs/01_通信录.xlsx"
    
    # 方法1：获取日期范围和考勤数据
    date_range, attendance_data = process_pc_attendance(pc_attendance_path) 

    # 初始化考勤数组
    contact_attendance_list = init_attendance_template(person_data, date_range[0], date_range[1])

    # 索引构建
    index_map = build_record_index(contact_attendance_list)

    # 写入PC考勤
    fill_pc_attendance(index_map, attendance_data)

    # 移动OA
    fill_oa_attendance(index_map, oa_data)

    fill_leave_registration(index_map, leave_df)

    fill_leave_info(index_map, qj_df)
    
    summary_result = summarize_attendance(contact_attendance_list, holiday_set)
    df_summary = pd.DataFrame(summary_result)
    
    output_path = "考勤统计结果.xlsx"
    df_summary.to_excel(output_path, index=False)

    print(f"✅ 成功导出到：{output_path}")

