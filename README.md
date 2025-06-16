### discord自動發言廣告
![version](https://img.shields.io/badge/version-1.0.0-blue)
![license](https://img.shields.io/badge/license-MIT-green)
![python](https://img.shields.io/badge/python-3.8+-yellow)
![GitHub issues](https://img.shields.io/github/issues/chase5ws/forward-tek_mail_automaticreply)
![GitHub stars](https://img.shields.io/github/stars/chase5ws/forward-tek_mail_automaticreply)
![GitHub forks](https://img.shields.io/github/forks/chase5ws/forward-tek_mail_automaticreply)
![icon](asset/icon.png)

---
* 使用main.exe檔案即可執行使用

---
### 專案簡介

此專案用於快速生成與處理離職同仁的資料，包含生成測試資料、處理現有資料並輸出結果，以及將資料轉換為文字檔案等功能。適合需要批次處理離職資料的場景，提升工作效率。

---
### 功能概述

- **生成測試資料**：快速生成範例 Excel 檔案，方便測試或模擬操作。
- **處理離職資料**：讀取 Excel 檔案，根據狀態自動生成標準化回覆內容，並輸出新檔案。
- **轉換文字檔案**：將每列 Excel 資料轉換為文字檔案，便於後續使用。

---
### 技術重點

- **pandas**：用於高效處理 Excel 資料。
- **openpyxl**：支援 Excel 檔案的讀寫操作。
- **os**：處理檔案與資料夾操作。
- **datetime**：生成時間戳，避免檔案覆蓋。
