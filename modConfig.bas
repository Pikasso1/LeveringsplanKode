Attribute VB_Name = "modConfig"
' ========= modConfig.bas =========
Option Explicit

' -------- Sheets --------
Public Const MASTER_SHEET_NAME As String = "Master"
Public Const LEVERINGSPLAN_PREFIX As String = "Leveringsplan "  ' e.g. "Leveringsplan 2025"

' -------- Master sheet fixed columns --------
' Adjust to your real layout. Example below assumes:
'   Col A = Varenr, Col B = Navn, Col C = HoursPerItem, Col D = PaintClass, Col E = ProdNote
Public Const MASTER_DATA_FIRST_ROW As Long = 2
Public Const MASTER_COL_NAVN       As Long = 1  ' A
Public Const MASTER_COL_VARENR     As Long = 4  ' D
Public Const MASTER_COL_PAINT      As Long = 9  ' I
Public Const MASTER_COL_HOURS      As Long = 10 ' J
Public Const MASTER_COL_PRODNOTE   As Long = 13 ' M

' Normalisation behaviour for varenr keys
Public Const MASTER_TRIM_KEYS       As Boolean = True
Public Const MASTER_UPPERCASE_KEYS  As Boolean = True   ' If True, force uppercase keys; if false, do nothing
Public Const MASTER_ALLOW_DUP_VNR   As Boolean = False  ' If False, first wins; duplicates are logged

' -------- Leveringsplan fixed columns (for Step 3 writer; declared now to keep a single source) --------
Public Const PLAN_COL_HEADER_A        As Long = 1   ' A
Public Const PLAN_COL_NAVN            As Long = 1   ' A  (data rows) / also holds headers
Public Const PLAN_COL_ANTAL           As Long = 2   ' B
Public Const PLAN_COL_ORDERNO         As Long = 3   ' C
Public Const PLAN_COL_VARENR          As Long = 4   ' D

Public Const PLAN_COL_GAP_1           As Long = 5   ' E (leave untouched)
Public Const PLAN_COL_GAP_2           As Long = 6   ' F (leave untouched)
Public Const PLAN_COL_DATO            As Long = 7   ' G
Public Const PLAN_COL_GAP_3           As Long = 8   ' H (leave untouched)

Public Const PLAN_COL_PAINT           As Long = 9   ' I
Public Const PLAN_COL_HOURS_PER_ITEM  As Long = 10  ' J   <-- changed semantics
Public Const PLAN_COL_GAP_4           As Long = 11  ' K (leave untouched)
Public Const PLAN_COL_TOTAL_HOURS     As Long = 12  ' L   <-- new explicit
Public Const PLAN_COL_PRODNOTE        As Long = 13  ' M

Public Const PLAN_COL_CLOSED          As Long = 3   ' C

Public Const CAT_PROD_MAL_SAME  As String = "Produceres, males og leveres i ugen"
Public Const CAT_PROD_SAME      As String = "Produceres og leveres i ugen"
Public Const CAT_PROD_THIS_NEXT As String = "Produceres i denne uge , males og leveres i næste"
Public Const CAT_Q_STOP         As String = "Cylindre med Q-stop  Produceres i denne uge , tjekkes og males i de to næste"
Public Const CAT_LAGER          As String = "Lagerproduktioner"
Public Const CAT_DIVERSE        As String = "Diverse til levering i ugen"

' --------- Not Planned overview config ---------
Public Const ERP_PLANNED_COLOR              As Long = 5296274   ' "Planlagt" grøn farve #50D092
Public Const ERP_PACKED_COLOR               As Long = 15773696  ' "Planlagt" blå farve  #00B0F0
Public Const PLAN_STATUS_COLOR_CHECK_COL    As Long = PLAN_COL_GAP_2    ' F
Public Const PLAN_COL_CAPACITY_TOTAL        As Long = PLAN_COL_DATO     ' G
Public Const PLAN_COL_CAPACITY_USED         As Long = PLAN_COL_GAP_4    ' K
Public Const PLAN_COL_PROTOTYPE             As Long = PLAN_COL_GAP_4    ' K
Public Const OVERVIEW_SHEET_NAME            As String = "Not Planned"
Public Const NOTPLANNED_START_YEAR          As Long = 2025
Public Const NOTPLANNED_START_WEEK          As Long = 23
Public Const NOTPLANNED_SKIP_6DIGIT_VARENR  As Boolean = True

' Layout constants for all gauges & forms
Public Const CFG_WEEK_LABEL_WIDTH       As Long = 60      ' px width of the “Uge” label
Public Const CFG_ROW_HEIGHT             As Long = 24     ' px height per week-row
Public Const CFG_BOTTOM_MARGIN          As Long = 6      ' px extra under the last row

' If your form preview uses different constants, put them here too:
Public Const CFG_PREVIEW_ROW_HEIGHT     As Long = 12    ' for lstPreview.Rows
Public Const CFG_FORM_BOTTOM_GAP        As Long = 40    ' for buttons under lstPreview
Public Const PREVIEW_ORIG_TOP           As Long = 200   ' What its set to in the designer
Public Const PREVIEW_ORIG_HEIGHT        As Long = 36    ' Original preview height
Public Const PREVIEW_GAUGE_MARGIN       As Long = 6     ' Preview.top - (Gauge.Top + height)
Public Const GAUGE_ORIG_TOP             As Long = 170   ' What its set to in the designer
Public Const GAUGE_ORIG_HEIGHT          As Long = 24    ' Original gauge parent frame height

'— Gauge bar fill colors
Public Const CFG_COLOR_USED_FILL        As Long = vbGreen
Public Const CFG_COLOR_ORDER_FILL       As Long = vbRed
Public Const CFG_COLOR_REMAIN_FILL      As Long = vbYellow

'— Overflow week-label background
Public Const CFG_COLOR_OVERFLOW_LABEL   As Long = vbRed
Public Const CFG_COLOR_DEFAULT_LABEL    As Long = vbButtonFace


