# ui/constants.py — 全域 UI 常數
# 所有面板 import 這裡的常數，保持視覺一致性

# ── Padding 系統 ──────────────────────────────────────
PX      = 14          # LabelFrame 水平邊距
PY_S    = (8, 4)      # 一般 section 間距 (top, bottom)
PY_B    = (8, 12)     # 最後一個 section 間距
PYW     = 5           # widget 上下間距

# ── 按鈕 bootstyle 語義 ───────────────────────────────
BTN_BROWSE    = "secondary-outline"   # 瀏覽…、開啟資料夾
BTN_UTILITY   = "secondary-outline"   # 重新整理、重新載入
BTN_ACTION    = "secondary-outline"   # 今天、搜尋等次要動作
BTN_PRIMARY   = "primary"             # 每個 section 的主要操作
BTN_ADD       = "primary-outline"     # 新增類操作
BTN_GENERATE  = "success"             # 產生報告
BTN_SAVE      = "success"             # 儲存
