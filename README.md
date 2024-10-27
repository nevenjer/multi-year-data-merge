import pandas as pd

# อ่านข้อมูลจากสามชีท
sheet_2565 = pd.read_excel(r"C:\Your Path File", sheet_name="Sheet1")  # ปี 2565
sheet_2566 = pd.read_excel(r"C:\Your Path File", sheet_name="Sheet2")  # ปี 2566
sheet_2567 = pd.read_excel(r"C:\Your Path File", sheet_name="Sheet3")  # ปี 2567

# รวมข้อมูล โดยใช้ acc_no และ acc_name เป็น key
combined = pd.merge(sheet_2565, sheet_2566, on=["acc_no", "acc_name"], how="outer", suffixes=('_2565', '_2566'))
combined = pd.merge(combined, sheet_2567, on=["acc_no", "acc_name"], how="outer", suffixes=('', '_2567'))

# เปลี่ยนชื่อคอลัมน์ให้ตรงตามที่ต้องการ
combined.rename(columns={
    'cost_2565': 'cost_2565',
    'plan_2565': 'plan_2565',
    'cost_2566': 'cost_2566',
    'plan_2566': 'plan_2566',
    'cost_2567': 'cost_2567',
    'plan_2567': 'plan_2567'
}, inplace=True)

# บันทึกข้อมูลรวมลงไฟล์ใหม่
combined.to_excel(r"C:\Your Path File\combined_data.xlsx", index=False)
print("-------------The combined data has been successfully completed!!"-------------)
