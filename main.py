import pandas as pd
import glob
import os
import re

def process_daily_reports(folder_path):
    print("📥 อ่านข้อมูล Daily Reports...")
    data = []

    for file_path in glob.glob(os.path.join(folder_path, "*.xls*")):
        try:
            df = pd.read_excel(file_path)

            filename = os.path.basename(file_path)
            match = re.search(r"Daily report_(\d+)_(.*)\.xls", filename)
            if not match:
                match = re.search(r"Daily report_(\d+)_(.*)\.xlsx", filename)
            if match:
                date_str = match.group(1)
                team_member = match.group(2).replace("_", " ")
            else:
                date_str = "0"
                team_member = "Unknown"

            if "Status" not in df.columns or "Candidate Name" not in df.columns:
                print(f"⚠️ ไฟล์ {filename} ไม่มีคอลัมน์ 'Status' หรือ 'Candidate Name'")
                continue

            passed = df[df["Status"] == "Pass"]

            for _, row in passed.iterrows():
                data.append({
                    "Employee Name": row["Candidate Name"],
                    "Team Member": team_member,
                    "Date": date_str
                })
        except Exception as e:
            print(f"❌ Error reading {file_path}: {e}")

    daily_df = pd.DataFrame(data)
    if not daily_df.empty:
        daily_df = daily_df.sort_values("Date").drop_duplicates(subset=["Employee Name"], keep="first")
        daily_df.drop(columns=["Date"], inplace=True)
    
    return daily_df

def process_new_employees(file_path):
    print("📥 อ่านข้อมูลพนักงานใหม่...")
    try:
        df = pd.read_excel(file_path)
        # ลบคอลัมน์ที่ไม่ต้องการ
        df = df.drop(columns=["DOB", "ID Card"], errors='ignore')
        if 'Join Date' in df.columns:
            df['Join Date'] = pd.to_datetime(df['Join Date']).dt.date
        return df
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return pd.DataFrame()

def merge_and_create_dashboard(daily_df, new_emp_df):
    print("🔄 รวมข้อมูลและสร้าง Dashboard...")

    if daily_df.empty or new_emp_df.empty:
        print("⚠️ ไม่พบข้อมูลที่ตรงกันระหว่างผู้ผ่านสัมภาษณ์กับพนักงานใหม่")
        return pd.DataFrame()

    # ล้างช่องว่าง
    daily_df["Employee Name"] = daily_df["Employee Name"].str.strip()
    new_emp_df["Employee Name"] = new_emp_df["Employee Name"].str.strip()

    merged = pd.merge(new_emp_df, daily_df, how='left', on="Employee Name")

    # 🔍 เติม Team Member = "Unknown" ถ้าไม่มีข้อมูล
    merged["Team Member"] = merged["Team Member"].fillna("Unknown")

    # 🔍 พิมพ์ debug
    print("🔍 ตัวอย่างข้อมูลหลัง merge:")
    print(merged[["Employee Name", "Team Member"]].head(10))

    final_dashboard = merged[["Employee Name", "Join Date", "Role", "Team Member"]]
    return final_dashboard


def main():
    daily_folder = "daily_reports"
    
    new_emp_file = "new_employees/New Employee_YYYYMM.xlsx"
    
    output_file = "dashboard_output.xlsx"

    daily_df = process_daily_reports(daily_folder)
    new_emp_df = process_new_employees(new_emp_file)

    dashboard = merge_and_create_dashboard(daily_df, new_emp_df)

    if not dashboard.empty:
        dashboard.to_excel(output_file, index=False)
        print(f"✅ สร้าง Dashboard เสร็จแล้ว: {output_file}")
    else:
        print("⚠️ ไม่มีข้อมูลให้สร้าง Dashboard")

if __name__ == "__main__":
    main()