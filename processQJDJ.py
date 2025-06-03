import pandas as pd
from datetime import timedelta

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