VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsLeveringsPlanForm 
   Caption         =   "Indsæt nye salgsordre linjer"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "clsLeveringsPlanForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "clsLeveringsPlanForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private pOrderLines As New cOrderLines          ' the staging bucket
Private dictKategori As Object                  ' Input -> kategori dictionary
Private WithEvents fraGauge As MSForms.Frame    ' ensure your frame is named fraGauge
Attribute fraGauge.VB_VarHelpID = -1


Private Function ParseList(ByVal txt As Variant, _
                           Optional keepBlanks As Boolean = False) As String()
    If IsNull(txt) Then txt = ""
    Dim bits, out() As String, i&, n&
    bits = Split(Replace(CStr(txt), ";", ","), ",")
    ReDim out(0)
    For i = LBound(bits) To UBound(bits)
        If keepBlanks Or Len(Trim$(bits(i))) > 0 Then
            ReDim Preserve out(0 To n)
            out(n) = Trim$(bits(i))
            n = n + 1
        End If
    Next
    ParseList = out
End Function


Private Sub RefreshPreview()
    Dim idx As Long, ol As cOrderLine
    lstPreview.Clear
    
    ' Header row
    lstPreview.AddItem "Navn"                    ' 0 Name
    lstPreview.List(idx, 1) = "Vr. Nr"           ' 1 Product ID
    lstPreview.List(idx, 2) = "Antal"            ' 2 Amount
    lstPreview.List(idx, 3) = "t / Stk."         ' 3 Hrs/item
    lstPreview.List(idx, 4) = "Uge"              ' 4 Week
    lstPreview.List(idx, 5) = "År"               ' 5 Year
    lstPreview.List(idx, 6) = "Lev. Dato"        ' 6 Leveringsdato
    lstPreview.List(idx, 7) = "Kategori"         ' 7 Category
    idx = idx + 1
    
    ' Data fill
    For Each ol In pOrderLines.Items
        lstPreview.AddItem ol.Navn                   ' 0 Name
        lstPreview.List(idx, 1) = ol.varenr          ' 1 Product ID
        lstPreview.List(idx, 2) = ol.antal           ' 2 Amount
        lstPreview.List(idx, 3) = ol.HoursPerItem    ' 3 Hrs/item
        lstPreview.List(idx, 4) = ol.uge             ' 4 Week
        lstPreview.List(idx, 5) = ol.år              ' 5 Year
        lstPreview.List(idx, 6) = ol.Dato            ' 6 Leveringsdato
        lstPreview.List(idx, 7) = ol.kategori        ' 7 Category Kategori
        idx = idx + 1
    Next

    Const MAX_ROWS As Long = 12                     ' show up to 12 rows, then stop
    Dim rowsVisible As Long
    rowsVisible = Application.WorksheetFunction.Min( _
                    Application.WorksheetFunction.Max(1, lstPreview.ListCount), _
                    MAX_ROWS)
    
    lstPreview.Height = rowsVisible * CFG_PREVIEW_ROW_HEIGHT
    Me.Height = lstPreview.Top + lstPreview.Height + CFG_FORM_BOTTOM_GAP
    
    lblStatus.Caption = pOrderLines.count & " line(s) staged."
    lblStatus.BackColor = vbGreen
End Sub

'===================== StageLinesFromForm =====================
Private Sub StageLinesFromForm()
    ' -- 0 · reset staging so edits overwrite, not append
    Set pOrderLines = New cOrderLines

    lblStatus.Caption = ""
    lblStatus.BackColor = vbRed            ' assume error until success

    ' -- 1 · required: Kategori must be chosen
    'If Len(Trim(rawKategori)) = 0 Then
    '    lblStatus.Caption = "Kategori cannot be empty.": Exit Sub
    'End If

    ' -- 2 · read CSV strings (spaces NOT allowed) --------------------
    Dim rawVnr As String, rawAntal As String, rawDate As String
    Dim rawWeek As String, rawYear As String, rawKategori As String
    Dim gaugeHeight As Long

    rawVnr = CStr(cmbVarenummer.Value)
    rawAntal = CStr(cmbAntal.Value)
    rawDate = CStr(cmbLeveringsdato.Value)
    rawWeek = CStr(cmbUge.Value)
    rawYear = CStr(cmbÅr.Value)
    rawKategori = LCase(CStr(cmbKategori.Value))

    If InStr(rawVnr, " ") Or InStr(rawAntal, " ") Or InStr(rawDate, " ") _
         Or InStr(rawWeek, " ") Or InStr(rawYear, " ") Or InStr(rawKategori, " ") Then
        lblStatus.Caption = "Spaces are not allowed in CSV fields.": Exit Sub
    End If

    Dim vArr As Variant: vArr = ParseList(rawVnr)
    Dim aArr As Variant: aArr = ParseList(rawAntal)
    Dim dArr As Variant: dArr = ParseList(rawDate, True)  ' keep blanks
    Dim wArr As Variant: wArr = ParseList(rawWeek, True)
    Dim yArr As Variant: yArr = ParseList(rawYear, True)
    Dim katArr As Variant: katArr = ParseList(rawKategori)

    Dim nLines As Long: nLines = UBound(vArr)               ' 0-based

    ' ---- 3 · length checks for Vnr / Antal --------------------------
    If nLines < 0 Or CStr(vArr(0)) = "" Then
        lblStatus.Caption = "Nothing to stage.": Exit Sub
    End If
    If UBound(aArr) <> nLines Then
        lblStatus.Caption = "Antal count =/= Varenr count.": Exit Sub
    End If
    If UBound(katArr) <> nLines Then
        lblStatus.Caption = "Kategori count =/= Varenr count.": Exit Sub
    End If

    ' ---- 4 · pad Date, Week, Year lists -----------------------------
    If UBound(dArr) < nLines Then ReDim Preserve dArr(0 To nLines)

    If UBound(wArr) = 0 Then
        ReDim Preserve wArr(0 To nLines)
        Dim j&: For j = 0 To nLines: wArr(j) = wArr(0): Next
    ElseIf UBound(wArr) <> nLines Then
        lblStatus.Caption = "Week count =/= Varenr count.": Exit Sub
    End If

    If UBound(yArr) = 0 Then
        ReDim Preserve yArr(0 To nLines)
        For j = 0 To nLines: yArr(j) = yArr(0): Next
    ElseIf UBound(yArr) <> nLines Then
        lblStatus.Caption = "Year count =/= Varenr count.": Exit Sub
    End If

    ' ---- 5 · loop over each line -----------------------------------
    Dim i&, qty#, mi As cMasterItem, ok As Boolean
    Dim ol As cOrderLine

    For i = 0 To nLines
        ' Antal must be positive whole number
        If Not IsWholeNumber(aArr(i)) Or CLng(aArr(i)) <= 0 Then
            lblStatus.Caption = "Antal '" & aArr(i) & "' invalid on line " & (i + 1) & ".": Exit Sub
        End If
        qty = CLng(aArr(i))

        ' Master lookup or placeholder
        If vArr(i) = Empty Then
            lblStatus.Caption = "Fill in varenummer": Exit Sub
        End If
        
        ok = modMasterCache.TryGetItem(vArr(i), mi)
        If Not ok Then
            ' MsgBox "Varenr " & vArr(i) & " findes ikke i Master – fortsætter med tomme felter.", vbInformation
            Set mi = modMasterCache.CreatePlaceholderItem(vArr(i))
            Set ol = pOrderLines.AddLineFromMaster(mi, qty)   ' bypass 2nd lookup
        Else
            Set ol = pOrderLines.AddLine(vArr(i), qty)
        End If

        ' Year basic sanity (2020-2100)
        If IsWholeNumber(yArr(i)) Then
            If CLng(yArr(i)) < 2019 Or CLng(yArr(i)) > 2100 Then
                lblStatus.Caption = "Year '" & yArr(i) & "' invalid on line " & (i + 1) & ". Must be between 2020 & 2100": Exit Sub
            End If
        Else
            lblStatus.Caption = "Fill in år": Exit Sub
        End If
        ol.år = CLng(yArr(i))

        ' Week 1-52 validation
        If IsWholeNumber(wArr(i)) Then
            If CLng(wArr(i)) < 1 Or CLng(wArr(i)) > 52 Then
                lblStatus.Caption = "Week '" & wArr(i) & "' invalid on line " & (i + 1) & ".": Exit Sub
            End If
        Else
            lblStatus.Caption = "Fill in uge": Exit Sub
        End If
        ol.uge = CLng(wArr(i))

        ' Kategori
        Dim katAlias As String, actualKategori As String

        katAlias = katArr(i)
        If Trim(katAlias) = "" Or IsEmpty(katAlias) Then
            lblStatus.Caption = "Fill in kategori": Exit Sub
        ElseIf dictKategori.Exists(katAlias) Then
            actualKategori = dictKategori(katAlias)
        Else
            lblStatus.Caption = "Unknown Kategori alias '" & katAlias & "' on line " & (i + 1) & ".": Exit Sub
        End If
        
        ol.kategori = actualKategori
        

        ' --- Delivery date (week-level guard only) ---
        If Len(Trim$(dArr(i))) > 0 Then
            If Not IsDate(dArr(i)) Then
                lblStatus.Caption = "Bad date '" & dArr(i) & "' on line " & (i + 1) & ".": Exit Sub
            End If
            
            Dim deliv As Date: deliv = CDate(dArr(i))
            Dim dw As Long, dy As Long
            modPlanner.GetIsoWeekYear deliv, dw, dy  ' delivery week/year (ISO)
        
            ' Compare (dy, dw) >= (ol.år, ol.uge)
            If (dy < ol.år) Or (dy = ol.år And dw < ol.uge) Then
                lblStatus.Caption = "Delivery date (" & Format$(deliv, "dd-mm-yyyy") & _
                    ") is before production week " & ol.uge & " " & ol.år & ".": Exit Sub
            End If
        
            ol.Dato = deliv
        ElseIf RequiresDato(actualKategori) Then
            lblStatus.Caption = "Kategori requires date on line " & (i + 1) & ".": Exit Sub
        End If

        
        ' TODO Dont lock but make it an option
        ' If the cylinder has to be painted, then it has to be in a painted category
        'If CategoryNeedsPaint(ol.paintClass, ol.kategori) Then
        '        lblStatus.Caption = "Cylinder requires paint, change category on line " & (i + 1) & ".": Exit Sub
        'End If


        ' Salgsordre
        ol.OrderNo = cmbSalgsordre.Value
    Next i

    ' ---- 6 · success ------------------------------------------------
    lblStatus.BackColor = vbGreen
    lblStatus.Caption = pOrderLines.count & " line(s) staged."
    
    gaugeHeight = modGauge.RenderContainer(Me.fraGauge, pOrderLines, Me)
    Me.lstPreview.Top = GAUGE_ORIG_TOP + PREVIEW_GAUGE_MARGIN + gaugeHeight
    RefreshPreview
End Sub

' ========= helpers =========
Private Function IsWholeNumber(v) As Boolean
    If IsNumeric(v) And Trim(v) <> "" And Not IsEmpty(v) Then
        IsWholeNumber = True
        Exit Function
    End If
    
    IsWholeNumber = False
End Function

' Old dato requirement function. Replaced by RequiresDato in config
' Put in config for easier management later down the line
'Private Function NeedDate(cat As String) As Boolean
'    NeedDate = InStr(1, cat, "leveres i næste", vbTextCompare) > 0
'End Function

'===================== CommitStagedLines =====================
Private Sub CommitStagedLines()
    Dim ol As cOrderLine, salgsordre As String
    
    ' Indlæser salgsordrenummer
    salgsordre = cmbSalgsordre.Value
    
    If lblStatus.BackColor = vbRed Then
        MsgBox "Something is wrong, please fix all errors before continuing"
    Else
        For Each ol In pOrderLines.Items
            Debug.Print "Navn: " & ol.Navn
            Debug.Print "Varenummer: " & ol.varenr
            Debug.Print "Antal: " & ol.antal
            Debug.Print "Timer pr. genstand: " & ol.HoursPerItem
            Debug.Print "Uge: " & ol.uge
            Debug.Print "År: " & ol.år
            Debug.Print "Kategori: " & ol.kategori
            Debug.Print " "
        Next
        
        modDeliveryWriter.CommitStagedLines pOrderLines, salgsordre
        Call Clear_User_Input
    End If
End Sub

Private Sub Clear_User_Input()
    ' Clear previous user inputs, ready form for new order
    Set pOrderLines = New cOrderLines

    cmbSalgsordre.Value = ""
    cmbAntal.Value = ""
    cmbVarenummer.Value = ""
    cmbÅr.Value = ""
    cmbUge.Value = ""
    cmbLeveringsdato.Value = ""
    cmbKategori.Value = ""
    
    ' Call RenderContainer to use the gauge pipeline for clearing the weeks
    Call modGauge.RenderContainer(Me.fraGauge, pOrderLines, Me)
    
    ' Resize parent gauge after renderpipeline set its height to 0
    Me.fraGauge.Height = GAUGE_ORIG_HEIGHT
    
    ' Preview clear
    Me.lstPreview.Clear
    
    ' Resize form and preview window
    Me.Height = Me.Height + PREVIEW_ORIG_HEIGHT - Me.lstPreview.Height
    Me.lstPreview.Height = PREVIEW_ORIG_HEIGHT
    
    ' Reset Error/Status box
    Me.lblStatus.Caption = ""
    Me.lblStatus.BackColor = vbWindowBackground
    
End Sub

Private Sub UserForm_Initialize()
    Set dictKategori = CreateObject("Scripting.Dictionary")
    modMasterCache.InitMasterCache

    ' Load the dropdowns with data from the "Dropdown" sheet
    Dim ws As Worksheet
    Dim i As Long
    
    ' Reference the "Dropdown" sheet
    Set ws = ThisWorkbook.Sheets("Dropdown")
    
    ' Load år data from column A
    i = 2 ' Start from row 2
    Do While ws.Cells(i, "A").Value <> ""
        cmbÅr.AddItem ws.Cells(i, "A").Value
        i = i + 1
    Loop
    
    ' Load uge data from column B
    i = 2 ' Start from row 2
    Do While ws.Cells(i, "B").Value <> ""
        cmbUge.AddItem ws.Cells(i, "B").Value
        i = i + 1
    Loop
    
    ' Load Kategori data from column C & D
    i = 2
    Do While ws.Cells(i, "D").Value <> ""
        Dim shortHand As String, actualCat As String
        shortHand = LCase(Trim(ws.Cells(i, "C").Value))  ' alias in column C
        actualCat = Trim(ws.Cells(i, "D").Value)  ' actual category in column D
        
        ' Populate dictionary and combobox with shorthand
        If Not dictKategori.Exists(shortHand) Then dictKategori.Add shortHand, actualCat
        cmbKategori.AddItem shortHand ' helpful display
        i = i + 1
    Loop
    
    ' Dictionary debug log
    Dim shortHandDebug As Variant
    For Each shortHandDebug In dictKategori
        Debug.Print "Key: " & shortHandDebug & " Value: " & dictKategori(shortHandDebug)
    Next
End Sub

Private Sub cmbVarenummer_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub cmbAntal_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub cmbÅr_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub cmbUge_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub cmbKategori_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub cmbLeveringsdato_AfterUpdate()
    StageLinesFromForm
End Sub

Private Sub btnSubmit_Click()
    Dim response As VbMsgBoxResult
    Dim closedWeek As cOrderLine, isWeekClosed As Boolean
    Dim paintCatNeeded As cOrderLine
    Dim over As Collection, item As Variant
    Dim msg As String

    ' Stage lines
    StageLinesFromForm

    ' Week closed message box prep
    Set closedWeek = isOneWeekClosed(pOrderLines)
    isWeekClosed = (closedWeek.uge <> 0)

    ' Paint class vs category prep
    Set paintCatNeeded = CheckCategoryNeedsPaint(pOrderLines)

    ' === Overflow prep (uses the same model the gauges use) ===
    Set over = GetOverflowTuples(pOrderLines)

    ' Validation & confirmations (short-circuit on No)
    If lblStatus.BackColor = vbRed Then
        MsgBox "Something is wrong, please fix all errors before continuing"
        Exit Sub
    End If

    If isWeekClosed Then
        response = MsgBox("Uge " & closedWeek.uge & " er lukket, er du sikker på du vil lægge " & _
                          closedWeek.varenr & " ind i denne uge?", vbYesNo + vbExclamation)
        If response = vbNo Then Exit Sub
    End If

    If over.count > 0 Then
        msg = "Indsætning af salgsordren vil overfylde:" & vbCrLf
        For Each item In over
            msg = msg & "• " & CStr(item(0)) & " med " & _
                        Format(item(1), "#,##0.##") & " timer" & vbCrLf
        Next
        msg = msg & vbCrLf & "Vil du fortsætte?"

        response = MsgBox(msg, vbYesNo + vbExclamation, "Kapacitetsadvarsel")
        If response = vbNo Then Exit Sub
    End If

    If paintCatNeeded.paintCatNeeded Then
        response = MsgBox(paintCatNeeded.varenr & " har overfladebehandlingen " & _
                          paintCatNeeded.paintClass & ". Er du sikker på du vil lægge den i den umalet kategori """ & _
                          paintCatNeeded.kategori & """?", vbYesNo + vbExclamation)
        If response = vbNo Then Exit Sub
    End If

    ' Commit
    CommitStagedLines
End Sub

