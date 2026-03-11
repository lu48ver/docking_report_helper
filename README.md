# 塢修報告產生器 (Dock Repair Report Generator)

從 Excel 工程清單自動產生 Word 格式的塢修日報，適用於船舶進塢維修期間的每日工作報告。

## 功能

- 讀取 Excel 工程規劃清單，依日期自動擷取當日工項
- 支援船廠工程 / 自修工程分類
- 自動產生 Word (.docx) 報告，含照片表格預留欄位
- 自動建立照片分類資料夾
- 內建 Excel 編輯器，可直接修改工程清單
- 支援天氣資訊自動帶入
- 多種表格版面配置（跨欄、單欄、雙欄）
- 支援機艙 / 甲板部門切換
- GUI 介面，操作簡單

## 截圖

<!-- TODO: 加入應用程式截圖 -->

## 安裝

### 環境需求

- Python 3.10+

### 安裝步驟

```bash
# Clone repo
git clone https://github.com/<your-username>/reporthelper.git
cd reporthelper

# 建立虛擬環境（建議）
python -m venv venv
source venv/bin/activate  # Linux/Mac
# 或 Windows:
# venv\Scripts\activate

# 安裝套件
pip install -r requirements.txt
```

## 使用方式

```bash
python main.py
```

或在 Windows 上雙擊 `啟動報告產生器.bat`。

### 基本流程

1. **設定** — 選擇 Excel 工程清單、設定進塢起始日期與報告日期
2. **Excel 編輯器** — 檢視 / 編輯工程項目
3. **產生報告** — 選擇模板與表格版面，點擊產生

## 專案結構

```
reporthelper/
├── main.py                 # 程式進入點
├── requirements.txt        # Python 套件相依
├── templates/              # Word 報告模板
│   └── default/
│       └── template.docx
├── config/                 # 設定管理
│   ├── config_manager.py
│   └── app_settings.py
├── core/                   # 核心邏輯
│   ├── excel_processor.py  # Excel 資料處理
│   ├── report_generator.py # Word 報告產生
│   ├── folder_creator.py   # 照片資料夾建立
│   ├── template_manager.py # 模板管理
│   └── weather_service.py  # 天氣資訊服務
├── ui/                     # GUI 介面
│   ├── main_window.py
│   ├── config_panel.py
│   ├── excel_editor_panel.py
│   └── report_output_panel.py
└── 啟動報告產生器.bat       # Windows 快速啟動
```

## 授權

本專案採用 [GNU General Public License v3.0](LICENSE) 授權。

你可以自由使用、修改和散布本軟體，但任何衍生作品也必須以相同授權條款開源。
詳見 [LICENSE](LICENSE) 檔案。
