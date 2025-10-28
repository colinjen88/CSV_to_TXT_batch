
# CSV 批次轉換工具 / CSV Batch Converter

**中文說明**

一個現代化的 Windows 圖形介面工具，支援多檔案 CSV 批次轉換，可自訂輸出資料夾、輸出格式（txt/tsv/csv）、分隔符號，並可指定要刪除的欄位。介面美觀、操作簡單，適合資料前處理、報表轉檔等需求。

**主要特色：**
- 多檔案批次轉換
- 支援 txt/tsv/csv 輸出格式
- 可自訂分隔符號（逗號、Tab、分號）
- 可指定要刪除的欄位
- 輸出資料夾自訂
- 介面現代化（ttkbootstrap）

---

**English Description**

A modern Windows GUI tool for batch converting multiple CSV files. Supports custom output folder, output format (txt/tsv/csv), delimiter selection, and column removal. Beautiful interface and easy to use—perfect for data preprocessing and report conversion.

**Features:**
- Batch convert multiple files
- Output as txt, tsv, or csv
- Custom delimiter (comma, tab, semicolon)
- Remove specified columns
- Custom output folder
- Modern UI (ttkbootstrap)

---

## 介紹

這是一個簡易的 Windows GUI 工具，讓你可以批次選取多個 CSV 檔案，並將其轉換為不同格式的檔案（如 TXT, TSV, CSV）。你可以在轉換時指定要刪除的欄位、選擇輸出資料夾、輸出格式與分隔符號。

## 功能

- 支援多檔案批次轉換
- 可自訂要刪除的欄位標題（多組，逗號分隔）
- 可自訂輸出資料夾
- 支援多種輸出格式（txt, tsv, csv）
- 支援多種分隔符號（逗號, Tab, 分號）
- 轉換後的檔案皆為 UTF-8 編碼
- 圖形化操作介面，無需命令列



## 功能特色

- 批次選取多個 CSV 檔案，一鍵轉換為 TXT（UTF-8 編碼）
- 可自訂輸出資料夾，若不存在會自動建立
- 可指定要刪除的欄位（輸入標題名稱，支援多組）
- 支援多種輸出格式與分隔符號
	- **輸出格式**：txt、tsv、csv
	- **分隔符號**：逗號(,)、Tab(\t)、分號(;) 
- 現代化專業 UI（ttkbootstrap）
- 支援 Windows 執行檔打包

## 支援的轉換格式

- **txt**：純文字檔，可自訂分隔符號
- **tsv**：Tab 分隔值檔案
- **csv**：逗號分隔值檔案

### 可選分隔符號

- 逗號 ( , )
- Tab (\t)
- 分號 ( ; )

---

## Features

- Batch select multiple CSV files and convert to TXT (UTF-8)
- Custom output folder (auto-created if missing)
- Remove columns by header name (multiple supported)
- Support multiple output formats and delimiters
	- **Output formats**: txt, tsv, csv
	- **Delimiters**: comma (,), tab (\t), semicolon (;) 
- Modern professional UI (ttkbootstrap)
- Windows executable packaging supported

## Supported Conversion Formats

- **txt**: Plain text file, customizable delimiter
- **tsv**: Tab-separated values file
- **csv**: Comma-separated values file

### Available Delimiters

- Comma ( , )
- Tab (\t)
- Semicolon ( ; )
	```powershell
	pip install pyinstaller
	```
2. 在專案目錄下執行：
	```powershell
	pyinstaller --noconsole --onefile csv_to_txt_batch.py
	```
3. 執行檔會產生在 `dist` 資料夾內，直接複製該檔案即可在 Windows 執行。

## 使用方式
1. 點選「選取 CSV 檔案」選擇多個要轉換的檔案
2. 在「要刪除的欄位標題」欄位輸入要刪除的欄位名稱（多組用逗號分隔）
3. 點選「開始轉換」
4. 轉換完成後，會在原檔案目錄產生同名 .txt 檔

## 擴充性建議
- 可加入自訂輸出資料夾功能
- 支援其他分隔符號（如 TSV）
- 支援不同編碼格式選擇
- 增加轉換進度條與錯誤報告
- 支援命令列批次操作

## 貢獻
歡迎 issue/PR，請遵循 GitHub 標準流程。

## 授權
MIT License
