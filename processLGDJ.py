from hmac import new
import pandas as pd
from datetime import timedelta

def fill_leave_registration(index_map, leave_df):
    leave_df.columns = leave_df.columns.str.strip()
    leave_df["离岗日期"] = pd.to_datetime(leave_df["离岗日期"])
    leave_df["返岗日期"] = pd.to_datetime(leave_df["返岗日期"])
    leave_df["人员编码"] = leave_df["人员编码"].astype(str).str.strip()

    for _, row in leave_df.iterrows():
        emp_id = str(row["人员编码"]).strip().zfill(8)
        start_date = row["离岗日期"].date()
        
        # 如果返岗日期为 NaT，则默认为离岗日期
        if pd.isna(row["返岗日期"]):
            end_date = start_date
        else:
            end_date = row["返岗日期"].date()

        current_date = start_date
        while current_date <= end_date:
            key = (emp_id, current_date)

            if key in index_map:
                index_map[key]["oa离岗登记"] = True
            # else:
                # print(f"❗离岗登记表: {row},未找到 key: {key}，请确认 index_map 中是否存在")
            current_date += timedelta(days=1)