import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta

def process_shift_attendance(shift_df, punch_dict, index_map):
    """
    处理倒班人员的出勤记录，并输出关键调试信息
    """
    import traceback
    shift_day_dict = {}

    print("🟢 开始处理倒班出勤")

    for idx, row in shift_df.iterrows():
        try:
            emp_id = str(row.get("工号", "")).strip().zfill(8)
            start_time = pd.to_datetime(row.get("上班时间", ""), errors="coerce")
            end_time = pd.to_datetime(row.get("下班时间", ""), errors="coerce")

            if not emp_id or pd.isna(start_time) or pd.isna(end_time):
                print(f"⚠️ 跳过第{idx}行，数据不完整: emp_id={emp_id}, start={start_time}, end={end_time}")
                continue

            date_key = start_time.date()
            end_date_key = end_time.date()
            key = (emp_id, date_key)

            shift_day_dict[key] = True
            if end_date_key != date_key:
                shift_day_dict[(emp_id, end_date_key)] = True

            # 合并打卡记录
            punches_today = punch_dict.get((emp_id, date_key), [])
            punches_next_day = punch_dict.get((emp_id, end_date_key), [])
            punch_times = sorted(punches_today + punches_next_day)

            in_start = start_time - timedelta(hours=2)
            in_end = start_time + timedelta(minutes=30)
            out_start = end_time - timedelta(minutes=30)
            out_end = end_time + timedelta(hours=4)

            has_valid_in = any(in_start <= t <= in_end for t in punch_times)
            has_valid_out = any(out_start <= t <= out_end for t in punch_times)

            if key not in index_map:
                index_map[key] = {}

            index_map[key]["加班时长"] = 0  # 初始化

            if has_valid_in:
                if key not in index_map:
                    index_map[key] = {}
                index_map[key]["倒班出勤"] = True

            if has_valid_out:
                end_date_key = end_time.date()
                end_key = (emp_id, end_date_key)
                if end_key not in index_map:
                    index_map[end_key] = {}
                index_map[end_key]["倒班出勤"] = True

        except Exception as e:
            print(f"❌ 错误发生在第{idx}行，员工ID={row.get('工号')}")
            traceback.print_exc()
            raise  # 或 return shift_day_dict 提前结束

    print("🟢 倒班出勤处理完毕")
    return shift_day_dict



def process_overtime_and_guesthouse(punch_dict, org_dict, index_map):
    """
    针对所有有打卡记录的员工，计算加班时长、招待所员工出勤时长
    """
    for key, punch_times in punch_dict.items():
        emp_id, date = key
        punch_times = sorted(punch_times)

        if not punch_times:
            continue

        earliest = punch_times[0]
        latest = punch_times[-1]
        org_name = org_dict.get(key, "")

        if key not in index_map:
            index_map[key] = {}

        if "招待所" in org_name:
            duration = latest - earliest
            if duration >= timedelta(hours=8):
                index_map[key]["pc出勤状态"] = "正常出勤"
            elif duration >= timedelta(hours=7):
                index_map[key]["pc出勤状态"] = "缺勤"
            else:
                index_map[key]["pc出勤状态"] = ""
        else:
            standard_end = datetime.combine(latest.date(), datetime.strptime("18:30", "%H:%M").time())
            overtime = latest - standard_end
            if overtime.total_seconds() > 0:
                index_map[key]["加班时长"] = int(overtime.total_seconds() / 3600)


def fill_shift_attendance(index_map, shift_df, record_df):
    """
    主函数：处理倒班出勤、加班时长与招待所正常出勤
    """
    shift_df.columns = shift_df.columns.str.strip()
    record_df.columns = record_df.columns.str.strip()

    # Step 1: 构建打卡字典和组织名称字典
    punch_dict = defaultdict(list)
    org_dict = {}
    print("开始构建打卡字典")
    for _, row in record_df.iterrows():
        emp_id = str(row["工号"]).strip().zfill(8)
        punch_time = pd.to_datetime(row["考勤时间"], errors="coerce")
        org_name = str(row.get("所属组织", "")).strip()

        if pd.notna(punch_time):
            key = (emp_id, punch_time.date())
            punch_dict[key].append(punch_time)

            if key not in org_dict:
                org_dict[key] = org_name
    print("打卡字典构建完成")

    # Step 2: 处理倒班员工的出勤判断
    shift_day_dict = process_shift_attendance(shift_df, punch_dict, index_map)
    print("倒班员工出勤已经完成")

    # Step 3: 针对所有员工统计加班/出勤
    process_overtime_and_guesthouse(punch_dict, org_dict, index_map)
    print("加班已经完成")

    return shift_day_dict
