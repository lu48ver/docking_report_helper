# ui/constants.py — 全域 UI 常數
# 所有面板 import 這裡的常數，保持視覺一致性

# ── Padding 系統 ──────────────────────────────────────
PX      = 16          # 頁面水平邊距
PY_S    = (16, 4)     # 一般 section 間距 (top, bottom)
PY_B    = (16, 16)    # 最後一個 section 間距
PYW     = 5           # widget 上下間距

# ── 字體 ─────────────────────────────────────────────
SECTION_FONT = ("", 9, "bold")    # Section 標題字體
HINT_FONT    = ("", 9)            # 提示文字字體
LOG_FONT     = ("Consolas", 9)    # 執行記錄字體

# ── Label 寬度 ────────────────────────────────────────
LBL_W   = 9           # 標準標籤欄寬（確保表單對齊）

# ── 按鈕 bootstyle 語義 ───────────────────────────────
BTN_BROWSE    = "secondary-outline"   # 瀏覽…、開啟資料夾
BTN_UTILITY   = "secondary-outline"   # 重新整理、重新載入
BTN_ACTION    = "secondary-outline"   # 今天、搜尋等次要動作
BTN_PRIMARY   = "primary"             # 每個 section 的主要操作
BTN_ADD       = "primary-outline"     # 新增類操作
BTN_GENERATE  = "success"             # 產生報告
BTN_SAVE      = "success"             # 儲存
