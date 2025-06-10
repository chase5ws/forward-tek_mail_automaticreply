import pandas as pd
from datetime import datetime

# 生成測試資料的函數
def generate_test_excel(file_name="測試用excel.xlsx"):
    """
    生成一個測試用的 Excel 檔案，包含標題行和 3 行測試資料
    :param file_name: 檔案名稱
    """
    # 獲取當前日期
    current_date = datetime.now().strftime("%Y/%m/%d")  # 格式為 YYYY/MM/DD

    # 定義標題行
    header = ["MAIL", "NAME", "離職日期", "狀態"]
    
    # 定義測試資料（3 行）
    data = [
        ["test1@example.com", "測試員工1", current_date, "離職"],
        ["test2@example.com", "測試員工2", current_date, "停用"],
        ["test3@example.com", "測試員工3", current_date, "轉移"]
    ]
    
    # 將資料轉換為 DataFrame
    df = pd.DataFrame(data, columns=header)
    
    # 將資料寫入 Excel 檔案
    df.to_excel(file_name, index=False)
    print(f"測試 Excel 檔案已生成：{file_name}")

# 呼叫函數生成測試檔案
generate_test_excel()
