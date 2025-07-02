import pandas as pd
from collections import defaultdict
from datetime import timedelta

def fill_shift_attendance(index_map, shift_df, record_df):
    """
    判断是否为倒班出勤正常，并更新 index_map 中的信息。
    如果员工在排班时间内有打卡记录，则视为出勤正常。
    
    :param index_map: (工号, 日期) -> record 字典
    :param shift_df: 倒班表 DataFrame，包含上班时间、下班时间等字段
    :param record_df: 打卡记录 DataFrame，包含考勤时间字段
    """
    shift_df.columns = shift_df.columns.str.strip()
    record_df.columns = record_df.columns.str.strip()

    # 构建打卡时间字典：{(工号, 日期) -> [时间列表]}
    punch_dict = defaultdict(list)
    for _, row in record_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        dt = pd.to_datetime(row["考勤时间"], errors="coerce")
        if pd.notna(dt):
            punch_dict[(emp_id, dt.date())].append(dt)

    # 遍历倒班记录
    for _, row in shift_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        start_time = pd.to_datetime(row["上班时间"], errors="coerce")
        end_time = pd.to_datetime(row["下班时间"], errors="coerce")

        if pd.isna(start_time) or pd.isna(end_time):
            continue

        is_cross_day = str(row.get("是否跨天", "")).strip() == "是"
        key_day = start_time.date()
        key = (emp_id, key_day)

        # 打卡容差时间
        start_range = start_time - timedelta(minutes=60)
        end_range = end_time + timedelta(minutes=60)

        # 获取打卡时间列表，支持跨天获取下一天打卡记录
        punch_times = punch_dict.get((emp_id, key_day), [])
        if is_cross_day:
            punch_times += punch_dict.get((emp_id, end_time.date()), [])

        # 是否有打卡落在范围内
        is_attended = any(start_range <= t <= end_range for t in punch_times)

        if is_attended and key in index_map:
            index_map[key]["倒班出勤"] = True
