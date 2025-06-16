import pandas as pd

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
                "oa请假信息": "",
                "oa请假类型":"",
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


def summarize_attendance(contact_attendance_list, holiday_set):
    summary_map = {}

    for record in contact_attendance_list:

        # ✅ 如果是节假日，跳过这一天
        if record["考勤日期"] in holiday_set:
            continue

        name = record.get("姓名")
        emp_id = str(record.get("工号")).strip().zfill(8)
        dept = record.get("部门")

        pc_status = record.get("pc出勤状态")
        oa_status = record.get("oa出勤状态")
        oa_leave = record.get("oa请假信息")
        oa_leave_type = record.get("oa请假类型")
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
                "病假":0,
                "事假":0,
                "年休假":0,
                "婚丧假":0,
                "探亲假":0,
                "产假":0,
                "陪产假":0,
                "育儿假":0,
                "出差":0,
                "其他":0,
                "备注":"",
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
            if oa_leave_type == "病假":
                stat["病假"] += 1
            elif oa_leave_type == "事假":
                stat["事假"] += 1
            elif oa_leave_type == "年休假":
                stat["年休假"] += 1
            elif oa_leave_type == "婚丧假":
                stat["婚丧假"] += 1
            elif oa_leave_type == "探亲假":
                stat["探亲假"] += 1
            elif oa_leave_type == "产假":
                stat["产假"] += 1
            elif oa_leave_type == "陪产假":
                stat["陪产假"] += 1
            elif oa_leave_type == "育儿假":
                stat["育儿假"] += 1
            elif oa_leave_type == "请假类型未知":
                stat["备注"] += "请假类型未知"
        elif is_pc_normal or is_oa_normal:
            stat["正常出勤天数"] += 1
        else:
            stat["缺勤天数"] += 1

    # 返回聚合后的列表（可用于生成表格/导出）
    return list(summary_map.values())