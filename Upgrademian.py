import pandas as pd
import streamlit as st
import io
import re

# ตั้งค่าหน้าเว็บ
st.set_page_config(layout="wide")
st.title("📊 Dashboard พนักงานใหม่และผู้ผ่านการสัมภาษณ์")
st.markdown("---")

# ฟังก์ชันเดิมสำหรับ Daily Reports    
def process_daily_reports(uploaded_files):
    """
    ประมวลผลไฟล์ Daily Reports ที่ผู้ใช้อัปโหลดเข้ามา
    """
    data = []
    
    # 📝 ตรวจสอบและประมวลผลไฟล์ที่อัปโหลด
    if not uploaded_files:
        st.warning("⚠️ กรุณาอัปโหลดไฟล์ Daily Report")
        return pd.DataFrame()

    for file in uploaded_files:
        try:
            # 💡 ใช้ io.BytesIO เพื่อให้ pandas อ่านไฟล์จาก Streamlit ได้โดยตรง
            df = pd.read_excel(io.BytesIO(file.getvalue()))
            
            filename = file.name
            # 🔍 ปรับแก้ regex ให้ยืดหยุ่นขึ้น
            match = re.search(r"Daily report_(\d+)_?([^.]*)\.xls", filename, re.IGNORECASE)
            if match:
                date_str = match.group(1)
                # 💡 ใช้ .strip() เพื่อลบช่องว่างที่ไม่ต้องการ
                team_member = match.group(2).replace("_", " ").strip()
            else:
                date_str = "0"
                team_member = "Unknown"

            if "Status" not in df.columns or "Candidate Name" not in df.columns:
                st.warning(f"⚠️ ไฟล์ {filename} ไม่มีคอลัมน์ 'Status' หรือ 'Candidate Name'")
                continue

            passed = df[df["Status"] == "Pass"]

            for _, row in passed.iterrows():
                data.append({
                    "Employee Name": row["Candidate Name"],
                    "Team Member": team_member,
                    "Date": date_str
                })
        except Exception as e:
            st.error(f"❌ Error reading {file.name}: {e}")

    daily_df = pd.DataFrame(data)
    if not daily_df.empty:
        daily_df = daily_df.sort_values("Date").drop_duplicates(subset=["Employee Name"], keep="first")
        daily_df.drop(columns=["Date"], inplace=True)
    
    return daily_df

# ฟังก์ชันเดิมสำหรับ New Employees
def process_new_employees(uploaded_file):
    """
    ประมวลผลไฟล์ New Employees ที่ผู้ใช้อัปโหลดเข้ามา
    """
    if not uploaded_file:
        st.warning("⚠️ กรุณาอัปโหลดไฟล์ New Employee")
        return pd.DataFrame()

    try:
        df = pd.read_excel(io.BytesIO(uploaded_file.getvalue()))
        df = df.drop(columns=["DOB", "ID Card"], errors='ignore')
        if 'Join Date' in df.columns:
            df['Join Date'] = pd.to_datetime(df['Join Date'], errors='coerce').dt.date
        return df
    except Exception as e:
        st.error(f"❌ Error reading {uploaded_file.name}: {e}")
        return pd.DataFrame()

# ฟังก์ชันรวมข้อมูล
def merge_and_create_dashboard(daily_df, new_emp_df):
    """
    รวมข้อมูลและสร้างตาราง Dashboard
    """
    if daily_df.empty or new_emp_df.empty:
        return pd.DataFrame()
    
    daily_df["Employee Name"] = daily_df["Employee Name"].astype(str).str.strip()
    new_emp_df["Employee Name"] = new_emp_df["Employee Name"].astype(str).str.strip()

    merged = pd.merge(new_emp_df, daily_df, how='left', on="Employee Name")
    merged["Team Member"] = merged["Team Member"].fillna("Unknown")

    final_dashboard = merged[["Employee Name", "Join Date", "Role", "Team Member"]]
    return final_dashboard

# --- UI (User Interface) ของ Streamlit ---

# 💻 ส่วนอัปโหลดไฟล์
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. อัปโหลด Daily Reports (ไฟล์ Excel หลายไฟล์ได้)")
    uploaded_daily_files = st.file_uploader(
        "เลือกไฟล์ 'Daily report_...'",
        type=["xls", "xlsx"],
        accept_multiple_files=True
    )

with col2:
    st.subheader("2. อัปโหลดไฟล์ New Employees (ไฟล์เดียว)")
    uploaded_new_emp_file = st.file_uploader(
        "เลือกไฟล์ 'New Employee_...'",
        type=["xls", "xlsx"]
    )

# 🚦 ปุ่มประมวลผล
if st.button("🚀 สร้าง Dashboard"):
    if uploaded_daily_files and uploaded_new_emp_file:
        # 🤖 ประมวลผลข้อมูล
        with st.spinner("🔄 กำลังประมวลผลข้อมูล..."):
            daily_df = process_daily_reports(uploaded_daily_files)
            new_emp_df = process_new_employees(uploaded_new_emp_file)
        
        if not daily_df.empty and not new_emp_df.empty:
            dashboard = merge_and_create_dashboard(daily_df, new_emp_df)
            
            if not dashboard.empty:
                st.success("✅ สร้าง Dashboard สำเร็จแล้ว!")
                st.markdown("---")
                st.subheader("📊 ตารางสรุป Dashboard")
                st.dataframe(dashboard, use_container_width=True) # แสดง DataFrame บนเว็บ
                
                # 📥 เพิ่มปุ่มดาวน์โหลด
                excel_buffer = io.BytesIO()
                dashboard.to_excel(excel_buffer, index=False, engine='xlsxwriter')
                st.download_button(
                    label="📥 ดาวน์โหลด Dashboard เป็นไฟล์ Excel",
                    data=excel_buffer.getvalue(),
                    file_name="dashboard_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("⚠️ ไม่สามารถสร้าง Dashboard ได้เนื่องจากไม่มีข้อมูลที่ตรงกัน")
        else:
            st.error("❌ กรุณาตรวจสอบว่าไฟล์ที่อัปโหลดถูกต้องและมีข้อมูลครบถ้วน")
    else:
        st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ส่วนก่อนกดปุ่ม 'สร้าง Dashboard'")