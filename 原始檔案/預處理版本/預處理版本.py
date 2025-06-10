import os

def process_file(input_path, output_folder):
    # 讀取檔案內容
    with open(input_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # 解析檔案資料，過濾空行和多餘空格
    lines = [line.strip() for line in content.splitlines() if line.strip()]  # 去掉空行和空格

    # 如果有效行數不足 3 行，標記為錯誤
    if len(lines) < 3:
        mark_error(input_path, output_folder, "檔案格式不正確，行數不足")
        return

    # 取得 email、中文名稱、離職日期
    email = lines[0]
    # 將 email 中 @ 後面的部分移除
    email = email.split('@')[0]
    name = lines[1]
    last_working_day = lines[2]

    # 檢查是否有缺失的必要資訊
    if not email or not name or not last_working_day:
        mark_error(input_path, output_folder, "檔案缺少必要資訊 (email, 中文名稱, 或離職日期)")
        return

    # 替換佔位符
    processed_content = content.replace("XXX", name, 1)  # 第 1 個 XXX 替換為中文名稱
    processed_content = processed_content.replace("XXX", email, 1)  # 第 2 個 XXX 替換為 email
    processed_content = processed_content.replace("XXX", name, 1)  # 第 3 個 XXX 替換為中文名稱
    processed_content = processed_content.replace("[離職日期]", last_working_day)  # 替換 [離職日期]
    processed_content = processed_content.replace("[Full Name of Former Employee]", email)  # 替換 [Full Name of Former Employee]
    processed_content = processed_content.replace("[Last Working Day]", last_working_day)  # 替換 [Last Working Day]

    # 驗證替換後的結果是否正確
    if "XXX" in processed_content or "[離職日期]" in processed_content or "[Full Name of Former Employee]" in processed_content or "[Last Working Day]" in processed_content:
        mark_error(input_path, output_folder, "替換後仍有未處理的佔位符")
        return

    # 確保輸出目錄存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 輸出處理後的內容
    output_path = os.path.join(output_folder, os.path.basename(input_path))
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(processed_content)
    print(f"檔案處理完成: {output_path}")


def mark_error(input_path, output_folder, error_message):
    # 取得檔案名稱
    filename = os.path.basename(input_path)
    error_filename = f"error_{filename}"
    error_path = os.path.join(output_folder, error_filename)

    # 確保輸出目錄存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 寫入錯誤檔案
    with open(input_path, 'r', encoding='utf-8') as file:
        original_content = file.read()

    with open(error_path, 'w', encoding='utf-8') as file:
        file.write(f"【Error: {error_message}】\n\n")
        file.write(original_content)

    print(f"檔案標記為錯誤: {error_path}")


def process_folder(input_folder, output_folder):
    if not os.path.exists(input_folder):
        print(f"輸入資料夾不存在: {input_folder}")
        return

    for filename in os.listdir(input_folder):
        input_path = os.path.join(input_folder, filename)
        if os.path.isfile(input_path):  # 確保是檔案
            process_file(input_path, output_folder)

    print(f"所有檔案已處理完成，結果輸出至資料夾: {output_folder}")


if __name__ == "__main__":
    # 獲取當前程式所在的目錄
    current_folder = os.path.dirname(os.path.abspath(__file__))
    
    # 定義輸入與輸出資料夾
    input_folder = os.path.join(current_folder, "input")
    output_folder = os.path.join(current_folder, "output")

    # 檢查並處理 input 資料夾中的檔案
    process_folder(input_folder, output_folder)
