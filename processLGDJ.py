import pandas as pd
from datetime import timedelta

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