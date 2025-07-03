import pandas as pd
from collections import defaultdict
from datetime import timedelta

def fill_shift_attendance(index_map, shift_df, record_df):
    """
    判断是否为倒班出勤正常，并更新 index_map 中的信息。
    支持跨天下班情况，自动提取第二天打卡记录。
    """
    shift_df.columns = shift_df.columns.str.strip()
    record_df.columns = record_df.columns.str.strip()

    # 构建 {(工号, 日期): [打卡时间列表]}
    shift_day_dict = {}
    punch_dict = defaultdict(list)
    for _, row in record_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        punch_time = pd.to_datetime(row["考勤时间"], errors="coerce")
        if pd.notna(punch_time):
            punch_dict[(emp_id, punch_time.date())].append(punch_time)

    # 遍历倒班记录
    for _, row in shift_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        start_time = pd.to_datetime(row["上班时间"], errors="coerce")
        end_time = pd.to_datetime(row["下班时间"], errors="coerce")

        if pd.isna(start_time) or pd.isna(end_time):
            continue

        date_key = start_time.date()
        next_day_key = (end_time + timedelta(hours=4)).date()  # 考虑下班+容差后的第二天

        key = (emp_id, date_key)
        shift_day_dict[key] = True

        # 打卡时间范围
        in_start = start_time - timedelta(hours=2)
        in_end = start_time + timedelta(minutes=30)
        out_start = end_time - timedelta(minutes=30)
        out_end = end_time + timedelta(hours=4)

        # 打卡记录合并（当前日期 + 下一日）
        punch_times = punch_dict.get((emp_id, date_key), []) + punch_dict.get((emp_id, next_day_key), [])

        # 判断是否有合法的上下班打卡
        has_valid_in = any(in_start <= t <= in_end for t in punch_times)
        has_valid_out = any(out_start <= t <= out_end for t in punch_times)

        if has_valid_in and has_valid_out and key in index_map:
            index_map[key]["倒班出勤"] = True
    return shift_day_dict