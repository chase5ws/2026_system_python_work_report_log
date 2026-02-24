# 工時記錄統計工具 (Work Record Statistics Tool)
![version](https://img.shields.io/badge/version-1.0.0-green)
![license](https://img.shields.io/badge/license-MIT-blue)
![python](https://img.shields.io/badge/Python-3.7%2B-orange)
![tkinter](https://img.shields.io/badge/tkinter-built--in-lightgrey)
![openpyxl](https://img.shields.io/badge/openpyxl-3.0%2B-red)
![matplotlib](https://img.shields.io/badge/matplotlib-3.5%2B-purple)
![platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-brightgreen)

一個基於 Python + Tkinter 開發的圖形介面工具，專用於解析Excel工時記錄數據、生成加權統計圖表，並匯出格式化的Excel報告，支援多Sheet分類統計、日期智能分組、標題自動合併等功能。

## 🌟 功能特點
- **多Sheet數據解析**：自動識別Excel中多個Sheet的工時記錄，支援 `更新進度、狀態、作業名稱、目前進度、附註描述` 欄位匹配
- **日期區間篩選**：可自定義查詢起始/結束日期，自動解析標準日期（2026-02-01）和4位數日期（0205）格式
- **加權統計功能**：支援為不同Sheet設置加權值（輸入整數自動除以10），生成加權後的工作記錄統計堆疊圖
- **智能內容分組**：自動拆分多分隔符內容（\n/|/;/、），按4位數日期分組排序，並標註附註前綴
- **格式化Excel匯出**：
  - 標題自動合併欄位（A-D列），解決擠壓問題
  - 統一11號字體、自動換行、合理欄寬調整
  - 嵌入統計圖表，支援日期由新到舊排序
  - 標題樣式美化（藍色背景、白色文字、粗體）
- **繁體中文介面**：適配台灣/香港等地區使用習慣，支援微軟正黑體顯示
- **自定義視窗圖標**：支援加載自定義 PNG/ICO 圖標（my_icon.png）
- **配置自動保存**：Sheet加權值自動保存到JSON配置文件，下次啟動自動加載

## 📋 環境需求
### 1. 依賴套件
```bash
# 安裝必要套件
pip install openpyxl matplotlib
```
> 備註：tkinter 通常隨 Python 預裝，若缺失可透過以下方式補裝：
> - Windows：重新安裝 Python 並勾選「Tcl/Tk and IDLE」
> - macOS：`brew install python-tk`
> - Linux (Ubuntu/Debian)：`sudo apt-get install python3-tk`

### 2. 系統相容性
- Windows 7/10/11
- macOS 10.14+
- Linux (Ubuntu 18.04+)
- Python 版本：3.7 及以上

## 🚀 使用方法
### 1. 準備測試數據（可選）
將測試數據（TSV格式）複製到Excel，建立4個Sheet：`usercase`、`inhouse`、`case1`、`case2`，每個Sheet包含以下欄位：
```
更新進度	狀態	作業名稱	目前進度	附註描述
```

### 2. 自定義配置（可選）
編輯程式頂部的全域設定參數，調整視窗尺寸、字體、配置文件路徑等：
```python
WINDOW_SIZE = "900x700"  # 自定義窗口尺寸
FONT_NAME = "Microsoft JhengHei"  # 介面顯示字體
WEIGHT_CONFIG_FILE = "sheet_weight_config.json"  # 加權配置文件
ICON_PATH = "my_icon.png"  # 自定義圖標路徑
```
> 若需使用自定義視窗圖標，請將 `my_icon.png` 放在程式同一目錄下

### 3. 執行程式
將程式程式碼儲存為 `work_report_tool.py`，透過以下命令執行：
```bash
python work_report_tool.py
```

### 4. 操作步驟
1. 點擊「瀏覽」按鈕，選擇要解析的Excel工時記錄檔案（.xlsx格式）；
2. 設置查詢日期區間（預設為過去15天至當天），格式為 `YYYY-MM-DD`；
3. 點擊「讀取」按鈕，程式自動加載Excel中所有有效Sheet；
4. 在Sheet設置區域：
   - 勾選需要匯出的Sheet（預設全選）
   - 輸入每個Sheet的加權值（整數，自動除以10，例如輸入3=實際0.3）
5. 點擊「更新圖表」按鈕，生成加權後的工作記錄統計堆疊圖；
6. 點擊「匯出Excel」按鈕，選擇儲存位置，生成格式化的工作報告Excel檔案。

## 📁 專案結構
```
.
├── work_report_tool.py    # 主程式檔案
├── my_icon.png            # 自定義視窗圖標（可選）
├── sheet_weight_config.json  # 加權配置文件（自動生成）
├── test_data.txt          # 測試數據文件（可選）
└── README.md              # 使用說明文件
```

## ⚙️ 自定義擴展
### 1. 新增欄位匹配規則
修改 `HEADER_MAPPING` 字典，新增欄位關鍵字匹配規則：
```python
HEADER_MAPPING = {
    "更新進度": ["更新進度", "更新日期", "日期", "date", "執行日期"],  # 新增自定義關鍵字
    "狀態": ["狀態", "處理人", "負責人", "人員", "user", "status"],
    # 其他欄位...
}
```

### 2. 調整Excel匯出格式
修改匯出Excel的欄寬、字體大小、標題樣式等：
```python
# 調整欄寬
ws.column_dimensions["A"].width = 15  # 更新日期欄寬
ws.column_dimensions["D"].width = 100  # 工作內容欄寬

# 調整標題字體大小
date_title_cell.font = Font(bold=True, size=16, name=FONT_NAME)  # 日期區間標題改為16號字
```

### 3. 新增統計圖表類型
修改 `update_chart` 方法，將堆疊條形圖替換為其他類型（如折線圖）：
```python
# 替換為折線圖
ax.plot(x_pos, counts, label=sheet, color=colors[i % len(colors)])
```

## ⚠️ 注意事項
1. 僅支援.xlsx格式檔案，不支援舊版.xls格式；
2. 4位數日期解析規則：前2位為月份（1-12），後2位為日期（1-31），不符合則視為無效；
3. Excel匯出時圖表區域預留22行空間，若圖表被覆蓋可調整 `current_row += 22` 的數值；
4. 確保Excel檔案路徑不含特殊字元（如全形符號、空格），避免讀取失敗；
5. OneDrive目錄可能存在同步延遲，建議將檔案放在本地目錄處理；
6. 若自定義圖標加載失敗，程式會在控制台輸出提示，但不影響核心功能使用；
7. 加權值輸入非數字時，程式自動替換為預設值0.1。

## 🐞 常見問題
| 問題現象 | 可能原因 | 解決方案 |
|----------|----------|----------|
| 程式開啟後介面亂碼 | 系統缺少繁體中文字體 | 安裝「微軟正黑體 (Microsoft JhengHei)」 |
| 讀取Excel提示「無有效數據」 | 欄位名稱不匹配或日期格式錯誤 | 確認Excel包含「更新進度、狀態、作業名稱、目前進度、附註描述」欄位，日期格式為YYYY-MM-DD |
| 匯出Excel後標題未合併 | 欄位數量不匹配 | 確認MAX_COLUMN變量設為4（A-D列），檢查merge_cells語法是否正確 |
| 統計圖表中文亂碼 | 缺少中文字體配置 | 確保程式中包含 `plt.rcParams["font.sans-serif"] = [FONT_NAME, "SimHei"]` |
| 加權值未保存 | 配置文件無寫入權限 | 將程式放在可寫入的目錄（如桌面），避免系統目錄 |
| 自定義圖標不顯示 | 圖標檔案不存在或路徑錯誤 | 確認 `my_icon.png` 放在程式同一目錄，或修改 `ICON_PATH` 變量為絕對路徑 |
| PyInstaller打包提示圖標找不到 | 使用PNG格式打包EXE | 將PNG轉換為ICO格式，修改打包命令：`pyinstaller -F -w -i my_icon.ico work_report_tool.py` |

## 📦 打包為EXE（Windows）
### 1. 安裝PyInstaller
```bash
pip install pyinstaller
```

### 2. 打包命令
```bash
# 基礎打包（無自定義圖標）
pyinstaller -F -w work_report_tool.py

# 帶自定義圖標打包（需先將PNG轉為ICO）
pyinstaller -F -w -i my_icon.ico work_report_tool.py
```
> 參數說明：
> - `-F`：打包成單個EXE文件
> - `-w`：無控制台窗口（GUI程式建議加）
> - `-i`：指定ICO格式圖標文件

## 📄 免責聲明
> 本專案僅供教學與個人使用，開發者不承擔因使用本工具導致的任何數據丟失、檔案損壞等損失責任。

---

### 版本更新說明（v1.0.0）
1. 實現Excel工時記錄多Sheet解析與日期區間篩選
2. 新增Sheet加權統計功能，支援配置自動保存
3. 生成加權堆疊條形圖，支援圖表嵌入Excel
4. 優化Excel匯出格式，實現標題自動合併、字體統一、自動換行
5. 支援4位數日期智能分組與內容排序
6. 新增自定義視窗圖標加載功能
7. 完善異常處理，支援繁體中文介面顯示