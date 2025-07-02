import streamlit as st
import pandas as pd
from io import BytesIO
import time

from processPCKQ import process_pc_attendance, fill_pc_attendance
from processYDKQ import fill_oa_attendance
from processQJDJ import fill_leave_info
from processLGDJ import fill_leave_registration
from processCCKQ import fill_business_trip
from processShift import fill_shift_attendance  # ✅ 倒班处理逻辑
from all import init_attendance_template, build_record_index, summarize_attendance

st.set_page_config(page_title="考勤分析工具", layout="wide")
st.title("📊 考勤 Excel 分析器")

# 上传区域（折叠）
with st.expander("📂 上传所需文件（点击展开）", expanded=True):
    col1, col2, col3 = st.columns(3)
    uploaded_files = {}

    with col1:
        uploaded_files["person"] = st.file_uploader("🧑‍💼 通信录", type=["xlsx"], key="p1")
        uploaded_files["oa"] = st.file_uploader("🏢 OA打卡", type=["xlsx"], key="p2")
        uploaded_files["trip"] = st.file_uploader("🧳 出差记录", type=["xlsx"], key="p7")

    with col2:
        uploaded_files["pc"] = st.file_uploader("💻 PC考勤", type=["xlsx", "csv"], key="p3")
        uploaded_files["leave"] = st.file_uploader("📝 离岗登记", type=["xlsx"], key="p4")
        uploaded_files["shift"] = st.file_uploader("🕐 倒班记录", type=["xlsx"], key="p8")  # ✅ 新增上传倒班记录

    with col3:
        uploaded_files["qj"] = st.file_uploader("📅 请假记录", type=["xlsx"], key="p5")
        uploaded_files["holiday"] = st.file_uploader("🎉 节假日", type=["xlsx"], key="p6")
        uploaded_files["record"] = st.file_uploader("🕐 打卡记录", type=["xlsx", "csv"], key="p9")

# 分析按钮
if st.button("🚀 开始分析"):

    required_keys = ["person", "oa", "pc", "leave", "qj", "holiday", "trip", "shift", "record"]
    if not all(uploaded_files.get(k) for k in required_keys):
        st.error("❌ 请确保上传了所有文件。")
    else:
        with st.spinner("🕐 正在分析考勤数据，请稍候..."):
            # 标记一个开始时间
            start_time = time.time()
            person_df = pd.read_excel(uploaded_files["person"], usecols=["姓名", "工号", "所在部门"])
            oa_df = pd.read_excel(uploaded_files["oa"])
            leave_df = pd.read_excel(uploaded_files["leave"])
            qj_df = pd.read_excel(uploaded_files["qj"])
            holiday_df = pd.read_excel(uploaded_files["holiday"])
            holiday_set = set(pd.to_datetime(holiday_df["日期"]).dt.date)
            trip_df = pd.read_excel(uploaded_files["trip"])
            shift_df = pd.read_excel(uploaded_files["shift"])
            if uploaded_files["record"].name.endswith('.xlsx'):
                record_df = pd.read_excel(uploaded_files["record"])
            else:
                record_df = pd.read_csv(uploaded_files["record"], encoding='gbk', parse_dates=["考勤时间"])
            
            # 处理考勤模板
            date_range, attendance_data = process_pc_attendance(uploaded_files["pc"])

            contact_attendance_list = init_attendance_template(person_df, date_range[0], date_range[1])
            index_map = build_record_index(contact_attendance_list)
            # 填充各类考勤记录
            st.info("正在处理文件：PC考勤结果")
            fill_pc_attendance(index_map, attendance_data)
            st.info("正在处理文件：oa考勤记录")
            fill_oa_attendance(index_map, oa_df)
            st.info("正在处理文件：离岗登记")
            fill_leave_registration(index_map, leave_df)
            st.info("正在处理文件：请假登记")
            fill_leave_info(index_map, qj_df)
            st.info("正在处理文件：出差登记")
            fill_business_trip(index_map, trip_df)
            st.info("正在处理文件：倒班记录")
            fill_shift_attendance(index_map, shift_df, record_df)
                

            # 汇总统计
            st.info("正在汇总")
            summary_result = summarize_attendance(contact_attendance_list, holiday_set)
            df_summary = pd.DataFrame(summary_result)
            df_all = pd.DataFrame(contact_attendance_list)
            st.info("汇总完毕")
            # 标记一个结束时间
            end_time = time.time()
            # 计算并打印总耗时
            total_time = end_time - start_time
            st.info(f"分析耗时：{total_time:.2f} 秒")
            st.info("正在写入excel文件...")
            # 下载缓存
            buffer1 = BytesIO()
            df_summary.to_excel(buffer1, index=False)
            st.session_state["summary_bytes"] = buffer1.getvalue()

            buffer2 = BytesIO()
            df_all.to_excel(buffer2, index=False)
            st.session_state["detail_bytes"] = buffer2.getvalue()

            st.session_state["analysis_done"] = True

            # 展现正在写入excel的状态，并且未写入完成时展现动态的效果
            

# 下载区域
if st.session_state.get("analysis_done"):
    st.success("✅ 分析完成，请选择下载内容：")

    st.download_button(
        label="📥 下载汇总结果",
        data=st.session_state["summary_bytes"],
        file_name="考勤统计结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="📥 下载明细结果",
        data=st.session_state["detail_bytes"],
        file_name="考勤总览.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
