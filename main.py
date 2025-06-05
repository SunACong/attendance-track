# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# 自定义模块导入
from processPCKQ import process_pc_attendance, fill_pc_attendance
from processYDKQ import fill_oa_attendance
from processQJDJ import fill_leave_info
from processLGDJ import fill_leave_registration
from all import init_attendance_template, build_record_index, summarize_attendance

st.set_page_config(page_title="考勤分析工具", layout="wide")

st.title("📊 考勤 Excel 分析器")

# 上传区域
with st.expander("📂 上传所需文件（点击展开）", expanded=True):
    col1, col2, col3 = st.columns(3)

    uploaded_files = {}
    with col1:
        st.markdown("### 🧑‍💼 通信录")
        uploaded_files["person"] = st.file_uploader("通信录", type=["xlsx"], key="p1")
        st.markdown("### 🏢 OA打卡")
        uploaded_files["oa"] = st.file_uploader("OA打卡记录", type=["xlsx"], key="p3")

    with col2:
        st.markdown("### 💻 PC考勤")
        uploaded_files["pc"] = st.file_uploader("PC考勤表", type=["xlsx"], key="p2")
        st.markdown("### 📝 离岗登记")
        uploaded_files["leave"] = st.file_uploader("离岗登记表", type=["xlsx"], key="p4")

    with col3:
        st.markdown("### 📅 请假记录")
        uploaded_files["qj"] = st.file_uploader("请假记录", type=["xlsx"], key="p5")
        st.markdown("### 🎉 节假日")
        uploaded_files["holiday"] = st.file_uploader("节假日表", type=["xlsx"], key="p6")

# 按钮触发
if st.button("🚀 开始分析"):

    if not all(uploaded_files.values()):
        st.error("❌ 请确保上传了所有文件。")
    else:
        with st.spinner("🕐 正在分析考勤数据，请稍候..."):

            # 读取上传数据
            person_df = pd.read_excel(uploaded_files["person"], usecols=["姓名", "工号", "所在部门"])
            oa_df = pd.read_excel(uploaded_files["oa"])
            leave_df = pd.read_excel(uploaded_files["leave"])
            qj_df = pd.read_excel(uploaded_files["qj"])
            holiday_df = pd.read_excel(uploaded_files["holiday"])
            holiday_set = set(pd.to_datetime(holiday_df["日期"]).dt.date)

            # 处理 PC 考勤
            date_range, attendance_data = process_pc_attendance(uploaded_files["pc"])
            contact_attendance_list = init_attendance_template(person_df, date_range[0], date_range[1])
            index_map = build_record_index(contact_attendance_list)

            # 数据填充
            fill_pc_attendance(index_map, attendance_data)
            fill_oa_attendance(index_map, oa_df)
            fill_leave_registration(index_map, leave_df)
            fill_leave_info(index_map, qj_df)

            # 汇总统计
            summary_result = summarize_attendance(contact_attendance_list, holiday_set)
            df_summary = pd.DataFrame(summary_result)
            df_all = pd.DataFrame(contact_attendance_list)

            # 导出按钮
            st.success("✅ 分析完成！请下载结果：")

            summary_excel = BytesIO()
            df_summary.to_excel(summary_excel, index=False)
            st.download_button("📥 下载汇总结果", summary_excel.getvalue(), file_name="考勤统计结果.xlsx")

            all_excel = BytesIO()
            df_all.to_excel(all_excel, index=False)
            st.download_button("📥 下载所有明细", all_excel.getvalue(), file_name="考勤总览.xlsx")
