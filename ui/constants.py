# ui/constants.py

# Spacing system
PAGE_PAD_X = 22
PAGE_PAD_Y = 22
SECTION_GAP = 22
FIELD_GAP_Y = 10
FIELD_GAP_X = 10
LABEL_INPUT_GAP = 8
SECTION_PAD = 18
COMPACT_SECTION_PAD = 14

# Legacy aliases kept for compatibility with older imports
PX = PAGE_PAD_X
PY_S = (SECTION_GAP, 0)
PY_B = (SECTION_GAP, PAGE_PAD_Y)
PYW = FIELD_GAP_Y // 2

# Typography
UI_FONT = "Microsoft JhengHei UI"
SECTION_FONT = (UI_FONT, 10, "bold")
SECTION_DESC_FONT = (UI_FONT, 9)
HINT_FONT = (UI_FONT, 9)
BODY_FONT = (UI_FONT, 10)
VALUE_FONT = (UI_FONT, 11, "bold")
HERO_FONT = (UI_FONT, 18, "bold")
LOG_FONT = ("Consolas", 9)

# Sizing
LBL_W = 10
SHORT_INPUT_W = 18
MEDIUM_INPUT_W = 24

# Palette
APP_BG = "#edf1f5"
SURFACE_BG = "#ffffff"
SURFACE_MUTED_BG = "#f7f9fc"
SURFACE_SUBTLE_BG = "#f3f5f8"
BORDER_COLOR = "#d7dde5"
TEXT_COLOR = "#243042"
TEXT_MUTED = "#6c7785"
ACCENT_BG = "#e9f1ff"
ACCENT_FG = "#1f4f9f"
SUCCESS_BG = "#e9f8ef"
SUCCESS_FG = "#1f6a43"

# Button roles
BTN_BROWSE = "secondary-outline"
BTN_UTILITY = "secondary-outline"
BTN_ACTION = "secondary-outline"
BTN_PRIMARY = "primary"
BTN_ADD = "primary-outline"
BTN_GENERATE = "success"
BTN_SAVE = "success"
