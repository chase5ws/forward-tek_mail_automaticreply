import os

def generate_test_files():
    # 獲取當前程式所在的目錄
    current_folder = os.path.dirname(os.path.abspath(__file__))
    
    # 定義輸出資料夾名稱
    output_folder = os.path.join(current_folder, "input")
    
    # 確保輸出資料夾存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 測試檔案內容
    test_files = {
        "test1.txt": """example@example.com
王小明
2025-06-10

您好 XXX，

感謝您的服務。您的 email 是 XXX。
我們確認您於 [離職日期] 離職。
祝您未來一切順利。

此致，
[Full Name of Former Employee]
""",
        "test2.txt": """example@example.com
王小明
""",
        "test3.txt": """

王小明
2025-06-10

您好 XXX，

感謝您的服務。您的 email 是 XXX。
我們確認您於 [離職日期] 離職。
祝您未來一切順利。

此致，
[Full Name of Former Employee]
""",
        "test4.txt": """example@example.com
王小明
2025-06-10

您好 XXX，

感謝您的服務。您的 email 是 XXX。
我們確認您於 [離職日期] 離職。
祝您未來一切順利。

此致，
[Full Name of Former Employee]
[Last Working Day]
""",
        "test5.txt": """
""",
        "test6.txt": """example@example.com
王小明
2025-06-10

您好 XXX，

感謝您的服務。您的 email 是 XXX。
我們確認您於 [離職日期] 離職。
祝您未來一切順利。

此致，
XXX
"""
    }

    # 生成檔案
    for filename, content in test_files.items():
        file_path = os.path.join(output_folder, filename)
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"已生成測試檔案: {file_path}")

if __name__ == "__main__":
    # 生成測試檔案
    generate_test_files()
