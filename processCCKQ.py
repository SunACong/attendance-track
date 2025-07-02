import pandas as pd
from datetime import timedelta

def fill_business_trip(index_map, trip_df):
    """
    根据出差记录更新 index_map 中的考勤数据：标记出差信息（为 True）
    :param index_map: (工号, 日期) -> record
    :param trip_df: 出差 DataFrame
    """
    # 转换日期格式，非日期值将被转换为 NaT
    trip_df["出差开始日期"] = pd.to_datetime(trip_df["出差开始日期"], errors="coerce")
    trip_df["出差结束日期"] = pd.to_datetime(trip_df["出差结束日期"], errors="coerce")

    for _, row in trip_df.iterrows():
        # emp_id被错误识别为数字后带了.0后缀
        try:
            emp_id = str(int(row["人员编号"])).strip().zfill(8)
        except (ValueError, TypeError):
            emp_id = str(row["人员编号"]).strip().zfill(8)
        location = row.get("出差地点", "未知地点")

        # ✅ 校验开始与结束日期
        if pd.isna(row["出差开始日期"]) or pd.isna(row["出差结束日期"]):
            # print(f"⚠️ 无效出差记录：工号 {emp_id}，缺少开始或结束日期，已跳过。")
            continue

        start_date = row["出差开始日期"].date()
        end_date = row["出差结束日期"].date()

        # 日期逻辑校验（可选）
        if end_date < start_date:
            # print(f"⚠️ 异常出差记录：工号 {emp_id} 的结束日期早于开始日期，已跳过。")
            continue

        current_date = start_date
        while current_date <= end_date:
            key = (emp_id, current_date)
            print(key)
            if key in index_map:
                record = index_map[key]
                record["oa出差信息"] = True
                record["oa出差地点"] = location
                # print(f"❗出差登记表：{emp_id} {current_date} 未找到 key，已跳过")
            current_date += timedelta(days=1)
