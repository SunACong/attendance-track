import pandas as pd
from datetime import timedelta

# 加载数据
shift_df = pd.read_excel("cs\\倒班表.xlsx")
record_df = pd.read_csv("cs\\考勤记录.csv", encoding='gbk', parse_dates=["考勤时间"])

# 预处理
shift_df["上班时间"] = pd.to_datetime(shift_df["上班时间"])
shift_df["下班时间"] = pd.to_datetime(shift_df["下班时间"])
record_df["工号"] = record_df["工号"].astype(str).str.strip()

# 打卡允许范围（单位：分钟）
buffer_minutes = 120

results = []

for _, row in shift_df.iterrows():
    emp_id = str(row["工号"]).strip()
    name = row["姓名"]
    start_time = row["上班时间"]
    end_time = row["下班时间"]

    # 设置打卡有效时间范围
    valid_start = start_time - timedelta(minutes=buffer_minutes)
    valid_end = end_time + timedelta(minutes=buffer_minutes)

    # 过滤有效时间段内的打卡记录
    records = record_df[
        (record_df["工号"] == emp_id) &
        (record_df["考勤时间"] >= valid_start) &
        (record_df["考勤时间"] <= valid_end)
    ]

    # 提取最早/最晚打卡时间
    if not records.empty:
        clock_in = records["考勤时间"].min()
        clock_out = records["考勤时间"].max()

        # 判断是否“正常”打卡（打卡时间靠近上下班）
        is_in_valid = clock_in <= start_time + timedelta(minutes=10)
        is_out_valid = clock_out >= end_time - timedelta(minutes=10)

        status = "正常" if is_in_valid and is_out_valid else "异常"
    else:
        clock_in = clock_out = None
        status = "未打卡"

    results.append({
        "姓名": name,
        "工号": emp_id,
        "上班时间": start_time,
        "下班时间": end_time,
        "最早打卡": clock_in,
        "最晚打卡": clock_out,
        "状态": status
    })

# 输出结果
df_result = pd.DataFrame(results)
df_result.to_excel("倒班考勤统计结果.xlsx", index=False)

print("✅ 分析完成，结果已导出到：倒班考勤统计结果.xlsx")
