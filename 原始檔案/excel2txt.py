import pandas as pd
from datetime import datetime
import os

# 讀取 Excel 檔案
file_name = "測試用excel.xlsx"  # 替換為你的檔案名稱
sheet_name = 0  # 如果有多個 sheet，可以指定名稱或索引
df = pd.read_excel(file_name, sheet_name=sheet_name, header=None)  # 確保讀取所有資料

# 確保第一行是標題行
df.iloc[0] = ["MAIL", "NAME", "離職日期", "狀態"]  # 固定第一行的標題內容
df.columns = df.iloc[0]  # 將第一行設為 DataFrame 的標題行
df = df[1:]  # 刪除原本的第一行，因為已經作為標題行

# 獲取當前日期和時間
current_time = datetime.now().strftime("%m%d_%H%M%S")  # 格式為 MMDD_HHMMSS

# 建立資料夾
folder_name = f"output_{current_time}"  # 資料夾名稱
if not os.path.exists(folder_name):  # 如果資料夾不存在，則建立
    os.makedirs(folder_name)

# 處理每一列的資料（跳過標題行）
for index, row in df.iterrows():
    # 確保第 3 欄存在，並將其固定為 "N/A"
    if len(row) >= 3:  # 如果該列有至少 3 欄
        row.iloc[2] = "N/A"
    
    # 建立對應的 txt 檔案，存放在資料夾內
    txt_file_name = f"{current_time}_{index}.txt"  # 每一列對應一個檔案
    file_path = os.path.join(folder_name, txt_file_name)  # 確保路徑組合正確
    with open(file_path, "w", encoding="utf-8") as file:
        for value in row:
            if pd.notna(value):  # 檢查是否為非空值
                file.write(f"{value}\n")  # 直接寫入內容
        file.write("\n")  # 每列結束添加一行空行

print(f"處理完成，所有列內容已儲存到資料夾：{folder_name}")
