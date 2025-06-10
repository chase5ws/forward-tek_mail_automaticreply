import os
import sys
import importlib

def main():
    while True:
        print("\n=== 主程式選單 ===")
        print("1. 執行 generate_excel.py")
        print("2. 執行 process_excel.py")
        print("3. 執行 excel2txt.py")  # 新增選項
        print("4. 離開程式")
        choice = input("請選擇要執行的功能（輸入數字）：")

        if choice == "1":
            # 檢查 generate_excel.py 是否存在
            if os.path.exists("generate_excel.py"):
                print("\n正在執行 generate_excel.py...")
                generate_excel = importlib.import_module("generate_excel")
                if hasattr(generate_excel, "main"):
                    generate_excel.main()
                else:
                    print("錯誤：generate_excel.py 中沒有找到 main() 函數！")
            else:
                print("錯誤：generate_excel.py 檔案不存在！")
        elif choice == "2":
            # 檢查 process_excel.py 是否存在
            if os.path.exists("process_excel.py"):
                print("\n正在執行 process_excel.py...")
                process_excel = importlib.import_module("process_excel")
                if hasattr(process_excel, "main"):
                    process_excel.main()
                else:
                    print("錯誤：process_excel.py 中沒有找到 main() 函數！")
            else:
                print("錯誤：process_excel.py 檔案不存在！")
        elif choice == "3":
            # 檢查 excel2txt.py 是否存在
            if os.path.exists("excel2txt.py"):
                print("\n正在執行 excel2txt.py...")
                excel2txt = importlib.import_module("excel2txt")
                if hasattr(excel2txt, "main"):
                    excel2txt.main()
                else:
                    print("錯誤：excel2txt.py 中沒有找到 main() 函數！")
            else:
                print("錯誤：excel2txt.py 檔案不存在！")
        elif choice == "4":
            print("程式結束，再見！")
            sys.exit(0)
        else:
            print("無效的選項，請重新輸入！")

if __name__ == "__main__":
    main()
