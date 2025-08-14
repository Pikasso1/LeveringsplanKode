Attribute VB_Name = "modDeliveryWriter"
' ========= modDeliveryWriter.bas =========
Option Explicit

'--- PUBLIC ENTRY POINT -------------------------------------------------------
' Commit all staged lines in one transaction.
' • OrderLines  = your populated cOrderLines instance
' • SalesOrderNo is written in column C
Public Sub CommitStagedLines(ByVal OrderLines As cOrderLines, ByVal SalesOrderNo As String)
    If OrderLines Is Nothing Or OrderLines.count = 0 Then
        MsgBox "Nothing to commit.", vbInformation: Exit Sub
    End If

    Dim errs As Collection
    If Not OrderLines.ValidateAll(errs) Then
        Dim itm As Variant, msg$: msg = "Cannot commit – fixes needed:" & vbCrLf
        For Each itm In errs: msg = msg & "• " & itm & vbCrLf: Next
        MsgBox msg, vbExclamation: Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ROLLBACK

    '-- group by (År|Uge|Kategori) --
    Dim grpKey As String, groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim ol As cOrderLine
    For Each ol In OrderLines.Items
        grpKey = ol.år & "|" & ol.uge & "|" & ol.kategori
        If Not groups.Exists(grpKey) Then
            Set groups(grpKey) = New Collection
        End If
        groups(grpKey).Add ol
    Next

    '-- iterate groups and write --
    Dim k As Variant, parts() As String
    For Each k In groups.Keys
        parts = Split(k, "|")
        CommitGroup CLng(parts(0)), CLng(parts(1)), parts(2), _
                    SalesOrderNo, groups(k)
    Next

    MsgBox "Commit complete: " & OrderLines.count & " line(s) written.", vbInformation
    GoTo FINALLY

ROLLBACK:
    MsgBox "Commit aborted: " & Err.Description, vbCritical
FINALLY:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


'--- INTERNAL: write one (År,Uge,Kategori) collection  ------------------------
Private Sub CommitGroup(ByVal år As Long, ByVal uge As Long, _
                         ByVal kategori As String, ByVal SalesOrderNo As String, _
                         ByVal lines As Collection)

    'target sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LEVERINGSPLAN_PREFIX & år)

    'locate anchor rows
    Dim wkRow&, catRow&, nextHdr&, insertAt&
    wkRow = FindWeekRow(ws, uge, år)
    If wkRow = 0 Then Err.Raise 7301, , "Week " & uge & "-" & år & " not found."

    catRow = FindCategoryRowUnderWeek(ws, wkRow, år, kategori)
    If catRow = 0 Then Err.Raise 7302, , "Category '" & kategori & "' missing under week " & uge

    nextHdr = FindNextHeaderRow(ws, catRow, år)
    If nextHdr = 0 Then nextHdr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1

    insertAt = nextHdr          'insert directly above next header
    Dim n&: n = lines.count

    ws.Rows(insertAt).Resize(n).Insert xlDown

    '-- build 2-D array (A..M = 1..13)  --
    Dim arr(): ReDim arr(1 To n, 1 To 13)

    Dim i&, ol As cOrderLine
    i = 1
    For Each ol In lines
        arr(i, PLAN_COL_NAVN) = ol.Navn
        arr(i, PLAN_COL_ANTAL) = ol.antal
        arr(i, PLAN_COL_ORDERNO) = SalesOrderNo
        arr(i, PLAN_COL_VARENR) = ol.varenr

        If RequiresDato(kategori) Then arr(i, PLAN_COL_DATO) = ol.Dato

        arr(i, PLAN_COL_PAINT) = ol.paintClass
        arr(i, PLAN_COL_HOURS_PER_ITEM) = ol.HoursPerItem
        arr(i, PLAN_COL_TOTAL_HOURS) = ol.TotalHours
        arr(i, PLAN_COL_PRODNOTE) = ol.ProdNote
        i = i + 1
    Next

    '-- block-write once, zero-touch gaps (they’re Empty) --
    ws.Cells(insertAt, 1).Resize(n, 13).Value = arr
    
    ' Compute column letters once
    Dim letterQty As String, letterHPI As String, letterSetup As String, rowNum As Long
    
    With ws
      letterQty = Split(.Cells(1, PLAN_COL_ANTAL).Address(True, False), "$")(0)
      letterHPI = Split(.Cells(1, PLAN_COL_HOURS_PER_ITEM).Address(True, False), "$")(0)
      letterSetup = Split(.Cells(1, PLAN_COL_GAP_4).Address(True, False), "$")(0)
    
      ' Overwrite Total-Hours cells with =J[row]*B[row]+K[row]+1.5
      For rowNum = insertAt To insertAt + n - 1
        .Cells(rowNum, PLAN_COL_TOTAL_HOURS).Formula = _
          "=" & letterHPI & rowNum & "*" & letterQty & rowNum & _
          "+" & letterSetup & rowNum & "+1.5"
      Next rowNum
    End With
    

    '-- determine a data-row to copy formats from (not a header) --
    Dim fmtSource As Long
    ' 1) Prefer the row that was *just above* the insert point *before* we inserted.
    '    That row is now at insertAt + n  after the row-shift.
    Dim candidate As Long: candidate = insertAt + n
    
    If candidate > ws.Rows.count Then
        ' shouldn’t happen, but guard
    ElseIf Not IsHeaderRow(ws, candidate, år) And _
           Not IsRowEmpty(ws, candidate) Then
        fmtSource = candidate
    Else
        ' 2) Fallback: look *upwards* for the first non-header, non-blank row
        fmtSource = FindNearestDataRowUp(ws, catRow, insertAt, år)
    End If
    
    ' If we found something, copy formats once
    If fmtSource > 0 Then
        ws.Rows(fmtSource).Copy
        ws.Rows(insertAt).Resize(n).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    
    ' --- ALWAYS force column F to pure white fill (no pattern, RGB(255,255,255)) ---
    With ws.Range(ws.Cells(insertAt, PLAN_STATUS_COLOR_CHECK_COL), _
                  ws.Cells(insertAt + n - 1, PLAN_STATUS_COLOR_CHECK_COL)).Interior
        .Pattern = xlSolid          ' ensure a solid fill
        .color = RGB(255, 255, 255) ' absolute white, theme-independent
    End With
    
    ' --- Force column K to pure white fill if it does appear in the master ---
    If ws.Cells(insertAt + n - 1, PLAN_COL_NAVN).Value <> "" Then
        With ws.Range(ws.Cells(insertAt, PLAN_COL_PROTOTYPE), _
                      ws.Cells(insertAt + n - 1, PLAN_COL_PROTOTYPE)).Interior
            .Pattern = xlSolid          ' ensure a solid fill
            .color = RGB(255, 255, 255) ' absolute white, theme-independent
        End With
    Else
        With ws.Range(ws.Cells(insertAt, PLAN_COL_PROTOTYPE), _
                      ws.Cells(insertAt + n - 1, PLAN_COL_PROTOTYPE)).Interior
            .Pattern = xlSolid
            .color = RGB(146, 208, 80)
        End With
    End If

End Sub

'-- is this row blank? (no data in key data columns) --
Private Function IsRowEmpty(ws As Worksheet, ByVal r As Long) As Boolean
    IsRowEmpty = _
        Len(Trim$(ws.Cells(r, PLAN_COL_VARENR).Text)) = 0 And _
        Len(Trim$(ws.Cells(r, PLAN_COL_NAVN).Text)) = 0
End Function

'-- does Column A hold a header on this row? --
Private Function IsHeaderRow(ws As Worksheet, ByVal r As Long, ByVal yearNum As Long) As Boolean
    Dim aVal$: aVal = Trim$(ws.Cells(r, PLAN_COL_HEADER_A).Text)
    IsHeaderRow = IsWeekHeaderValue(aVal, yearNum) Or IsCategoryHeaderValue(aVal)
End Function

'-- walk upwards until we hit a data row or catRow (stop) --
Private Function FindNearestDataRowUp(ws As Worksheet, ByVal catRow As Long, _
                                      ByVal stopRow As Long, ByVal yearNum As Long) As Long
    Dim r As Long
    For r = stopRow - 1 To catRow + 1 Step -1
        If Not IsHeaderRow(ws, r, yearNum) And Not IsRowEmpty(ws, r) Then
            FindNearestDataRowUp = r
            Exit Function
        End If
    Next
End Function


