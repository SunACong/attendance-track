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

    person_dept_dict = dict(zip(unique_people["工号"], unique_people["所在部门"]))

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
                "oa请假天数": 0,
                "oa出差信息": "",
                "oa出差地点": "",
                "倒班出勤": "",
                "加班时长": 0,
                "是否异常": "",
            })
    return template_records, person_dept_dict


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

def summarize_attendance(contact_attendance_list, holiday_set, shift_day_dict):
    emp_shift_days = deal_shift(shift_day_dict)
    summary_map = {}
    for record in contact_attendance_list:
        emp_id = str(record.get("工号")).strip().zfill(8)
        attend_date = record["考勤日期"]
        oa_leave = record.get("oa请假信息")
        has_oa_leave = oa_leave is True
        

        name = record.get("姓名")
        dept = record.get("部门")

        pc_status = record.get("pc出勤状态")
        oa_status = record.get("oa出勤状态")
        oa_leave = record.get("oa请假信息")
        oa_leave_type = record.get("oa请假类型")
        oa_leave_days = record.get("oa请假天数")
        oa_absence = record.get("oa离岗登记")
        oa_clock = record.get("oa是否打卡")
        oa_trip = record.get("oa出差信息")
        shift_attended = record.get("倒班出勤")


        if emp_id not in summary_map:
            summary_map[emp_id] = {
                "姓名": name,
                "工号": emp_id,
                "部门": dept,
                "正常出勤天数": 0,
                "出差": 0,
                "迟到": 0,
                "早退": 0,
                "缺勤": 0,
                "旷工天数": 0,
                "病假": 0,
                "事假": 0,
                "年休假": 0,
                "婚丧假": 0,
                "探亲假": 0,
                "护理假": 0,
                "产假": 0,
                "陪产假": 0,
                "育儿假": 0,
                "未知请假类型": 0,
                "加班时长": 0, 
                "节假日打卡天数": 0,
                "旷工/请假天数": 0,
            }

        stat = summary_map[emp_id]
        
        is_all_empty = not pc_status and not oa_status and not oa_absence and not oa_leave and not oa_clock and not oa_trip and not shift_attended


        is_pc_normal = oa_absence is True or pc_status == "正常出勤"
        is_oa_normal = oa_status == "正常出勤"
        
        has_oa_trip = oa_trip is True
        is_shift_normal = shift_attended is True  # ✅ 倒班出勤判断

        # 获取员工倒班天数，如果工号不在emp_shift_days中，则默认为0
        total_shift_days = emp_shift_days.get(emp_id, 0)
        

        # 如果考勤日期是节假日且没有OA请假记录，则跳过当前记录
        if attend_date in holiday_set:
            if record.get("加班时长", 0) > 0:
                stat["节假日打卡天数"] += 1
            if not has_oa_leave and total_shift_days < 9:
                continue
        

        
        if is_all_empty and total_shift_days < 9:
            stat["旷工天数"] += 1
            stat["旷工/请假天数"] += 1
            record["是否异常"] = "是"
        elif has_oa_trip:
            stat["出差"] += 1
        elif has_oa_leave:
            stat["正常出勤天数"] += 1 - oa_leave_days
            stat["旷工/请假天数"] += 1
            if "病假" in oa_leave_type:
                stat["病假"] += oa_leave_days
            elif "事假" in oa_leave_type:
                stat["事假"] += oa_leave_days
            elif "年休假" in oa_leave_type:
                stat["年休假"] += oa_leave_days
            elif "婚" in oa_leave_type or "丧" in oa_leave_type:
                stat["婚丧假"] += oa_leave_days
            elif "探亲假" in oa_leave_type:
                stat["探亲假"] += oa_leave_days
            elif "产假" in oa_leave_type:
                stat["产假"] += oa_leave_days
            elif "陪产假" in oa_leave_type:
                stat["陪产假"] += oa_leave_days
            elif "护理假" in oa_leave_type:
                stat["护理假"] += oa_leave_days
            elif "育儿假" in oa_leave_type:
                stat["育儿假"] += oa_leave_days
            else:
                stat["未知请假类型"] += oa_leave_days
        elif is_pc_normal or is_oa_normal or is_shift_normal:
            stat["正常出勤天数"] += 1
        else:
            if total_shift_days < 9:
                if "迟到" in pc_status:
                    stat["迟到"] += 1
                elif "早退" in pc_status:
                    stat["早退"] += 1
                else:
                    stat["缺勤"] += 1
                record["是否异常"] = "是"
        stat["加班时长"] += record.get("加班时长", 0)
        if emp_id == "11990062":
                print(f"工号11990062 - {attend_date}: {stat}")
    return list(summary_map.values())

# 处理倒班出勤字典，字典的key是由工号和日期组成的元组
def deal_shift(shift_day_dict):

    # 创建一个字典来存储每个员工的倒班天数
    emp_shift_days = {}
    
    # 遍历shift_day_dict中的所有键
    for emp_id, date in shift_day_dict.keys():
        # 如果员工ID不在emp_shift_days中，初始化为0
        if emp_id not in emp_shift_days:
            emp_shift_days[emp_id] = 0
        # 增加该员工的倒班天数
        emp_shift_days[emp_id] += 1
    
    return emp_shift_days