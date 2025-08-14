Attribute VB_Name = "modNotPlannedOverview"
' ========= modNotPlannedOverview.bas =========
Option Explicit

Private count As Long

' === NEW: helper to skip 6-digit product IDs ===
Private Function ShouldSkipProductId(ByVal varenr As Variant) As Boolean
    Dim s As String
    s = Trim$(CStr(varenr))
    ' True if exactly six digits (e.g., "012345" is allowed/considered six digits)
    ShouldSkipProductId = (Len(s) = 6 And s Like "######")
End Function

' === NEW: helper to check if (year, week) is on/after the configured start ===
Private Function IsOnOrAfterStart(ByVal yearNum As Long, ByVal weekNum As Long) As Boolean
    If yearNum > NOTPLANNED_START_YEAR Then
        IsOnOrAfterStart = True
    ElseIf yearNum = NOTPLANNED_START_YEAR Then
        IsOnOrAfterStart = (weekNum >= NOTPLANNED_START_WEEK)
    Else
        IsOnOrAfterStart = False
    End If
End Function


' Finds the next category row below the given row
Private Function FindNextCategoryRow(ws As Worksheet, ByVal startRow As Long, ByVal yearNum As Long) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, PLAN_COL_HEADER_A).End(xlUp).Row
    For r = startRow To lastRow
        Dim val As String: val = Trim(ws.Cells(r, PLAN_COL_HEADER_A).Value)
        If IsCategoryHeaderValue(val) Then
            FindNextCategoryRow = r
            Exit Function
        ElseIf IsWeekHeaderValue(val, yearNum) Then
            Exit Function ' Next week reached
        End If
    Next r
End Function


' Safe-create sheet
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetOrCreateSheet.name = sheetName
    End If
End Function

' Extract the year from sheet name, assuming format "Leveringsplan YYYY"
Private Function ExtractYearFromSheetName(sheetName As String) As Long
    ExtractYearFromSheetName = CLng(Trim(Replace(sheetName, LEVERINGSPLAN_PREFIX, "")))
End Function

' Collect unplanned lines from a single sheet into the given collection
Private Sub CollectNotPlannedFromSheet(wsPlan As Worksheet, ByVal yearNum As Long, ByRef OrderLines As cOrderLines)
    Dim lastRow As Long, weekRow As Long, categoryRow As Long
    lastRow = wsPlan.Cells(wsPlan.Rows.count, PLAN_COL_HEADER_A).End(xlUp).Row

    weekRow = 1
    Do While weekRow <= lastRow
        Dim headerVal As String
        headerVal = Trim(wsPlan.Cells(weekRow, PLAN_COL_HEADER_A).Value)

        If IsWeekHeaderValue(headerVal, yearNum) Then
            Dim weekNum As Long: weekNum = CLng(Split(Split(headerVal, " ")(1), "-")(0))

            ' === NEW: only process weeks on/after NOTPLANNED_START_YEAR/NOTPLANNED_START_WEEK ===
            If IsOnOrAfterStart(yearNum, weekNum) Then

                ' Iterate categories within this week
                categoryRow = weekRow + 1
                Do
                    categoryRow = FindNextCategoryRow(wsPlan, categoryRow, yearNum)
                    If categoryRow = 0 Or categoryRow > lastRow Then Exit Do

                    Dim nextHeaderRow As Long
                    nextHeaderRow = FindNextHeaderRow(wsPlan, categoryRow, yearNum)
                    If nextHeaderRow = 0 Then nextHeaderRow = lastRow + 1

                    ' Scan each order line under this category
                    Dim dataRow As Long
                    For dataRow = categoryRow + 1 To nextHeaderRow - 1
                        If Len(Trim(wsPlan.Cells(dataRow, PLAN_COL_VARENR))) = 0 Then GoTo NextLine
                        
                        If wsPlan.Cells(dataRow, PLAN_STATUS_COLOR_CHECK_COL).Interior.color = ERP_PACKED_COLOR Then
                            Debug.Print "We hit a packed item" & wsPlan.Cells(dataRow, PLAN_COL_VARENR).Value
                        End If

                        If IsColorNotPlanned(wsPlan.Cells(dataRow, PLAN_STATUS_COLOR_CHECK_COL).Interior.color) Then
                            ' Unplanned line detected
                            Dim varenr As String: varenr = wsPlan.Cells(dataRow, PLAN_COL_VARENR).Value

                            ' === NEW: skip logging if varenr is exactly 6 digits ===
                            If NOTPLANNED_SKIP_6DIGIT_VARENR And ShouldSkipProductId(varenr) Then GoTo NextLine

                            Dim antal As Double: antal = wsPlan.Cells(dataRow, PLAN_COL_ANTAL).Value
                            Dim OrderNo As String: OrderNo = wsPlan.Cells(dataRow, PLAN_COL_ORDERNO).Value
                            Dim category As String: category = wsPlan.Cells(categoryRow, PLAN_COL_HEADER_A).Value
                            
                            Dim ol As cOrderLine
                            Set ol = OrderLines.AddLine(varenr, antal, dataRow)
                            ol.år = yearNum
                            ol.uge = weekNum
                            ol.kategori = category
                            ol.OrderNo = OrderNo
                            If RequiresDato(category) Then
                                ol.Dato = wsPlan.Cells(dataRow, PLAN_COL_DATO).Value
                            End If
                            ' Remember source sheet name for hyperlinks
                            ol.SourceSheet = wsPlan.name
                            
                            ' Raise the count thats used for msgbox to user
                            count = count + 1
                        End If
NextLine:
                    Next dataRow
                    categoryRow = nextHeaderRow
                Loop

            End If ' end IsOnOrAfterStart(year, week)
        End If
        weekRow = weekRow + 1
    Loop
End Sub

' Write overview hyperlinks using the stored SourceRow and sheet name
Private Sub WriteCombinedOverview(OrderLines As cOrderLines, wsOverview As Worksheet)
    Dim i As Long: i = 2
    Dim ol As cOrderLine
    For Each ol In OrderLines.Items
        wsOverview.Cells(i, 1).Value = ol.år
        wsOverview.Cells(i, 2).Value = ol.uge
        wsOverview.Cells(i, 3).Value = ol.OrderNo
        wsOverview.Cells(i, 4).Value = ol.varenr

        If ol.SourceRow <> 0 Then
            wsOverview.Hyperlinks.Add Anchor:=wsOverview.Cells(i, 5), _
                Address:="", _
                SubAddress:="'" & ol.SourceSheet & "'!A" & ol.SourceRow, _
                TextToDisplay:="Go to line"
        Else
            wsOverview.Cells(i, 5).Value = "Line not found"
        End If
        i = i + 1
    Next
End Sub

Private Function IsColorNotPlanned(ByVal color As Long) As Boolean
    ' Start as False
    IsColorNotPlanned = False
    
    ' If color doesnt matches agreed color scheme, then it is not planned
    If color <> ERP_PLANNED_COLOR And color <> ERP_PACKED_COLOR Then
        IsColorNotPlanned = True
    End If
End Function

Public Sub GenerateNotPlanned()
    Dim wsOverview As Worksheet
    Set wsOverview = GetOrCreateSheet(OVERVIEW_SHEET_NAME)
    wsOverview.Cells.ClearContents
    wsOverview.Range("A1:E1").Value = Array("Year", "Week", "OrderNo", "Varenr", "Link")
    
    RefreshMasterProductList
    InitMasterCache
    

    Dim ws As Worksheet
    Dim OrderLines As New cOrderLines
    
    ' Reset count for every time NotPlanned sheet is refreshed
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.name, Len(LEVERINGSPLAN_PREFIX)) = LEVERINGSPLAN_PREFIX Then
            Dim yearNum As Long
            yearNum = ExtractYearFromSheetName(ws.name)

            ' === NEW: skip entire sheets before NOTPLANNED_START_YEAR ===
            If yearNum >= NOTPLANNED_START_YEAR Then
                ' Collect all unplanned lines from this sheet (per-week gating happens inside)
                CollectNotPlannedFromSheet ws, yearNum, OrderLines
            End If
        End If
    Next ws

    ' After collecting from all sheets, write to overview
    WriteCombinedOverview OrderLines, wsOverview
    MsgBox "Finished processing: " & count & " unplanned lines (from week " & NOTPLANNED_START_WEEK & " in " & NOTPLANNED_START_YEAR & " and onwards)"
End Sub


