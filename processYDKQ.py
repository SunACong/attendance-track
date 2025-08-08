import pandas as pd


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
    oa_df["编号"] = oa_df["编号"].astype(str).str.zfill(8)
    grouped = oa_df.groupby(["编号", "打卡日期"])

    for (emp_id, date), group in grouped:
        has_morning = any(t < 9 for t in group["打卡小时"])
        has_evening = any(t > 18 or (t == 18 and m > 0) for t, m in zip(group["打卡小时"], group["打卡分钟"]))

        key = (emp_id, date)
        if key in index_map:
            record = index_map[key]
            record["oa出勤状态"] = "正常出勤" if has_morning and has_evening else "异常"
            record["oa是否打卡"] = True
        # else:
            # print(f"❗OA考勤表: {grouped},未找到 key: {key}，请确认 index_map 中是否存在") 