Attribute VB_Name = "modPlanner"
' ========= modPlanner.bas =========
Option Explicit

' --------------- PUBLIC API ---------------

' Return the row where Column A equals "<week>-<year>" (e.g. "34-2025")
Public Function FindWeekRow(ws As Worksheet, _
                             ByVal weekNum As Long, ByVal yearNum As Long) As Long
    Dim key As String: key = "Uge " & CStr(weekNum) & "-" & CStr(yearNum)

    Dim hit As Range
    Set hit = ws.Columns(PLAN_COL_HEADER_A).Find(what:=key, _
                LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not hit Is Nothing Then FindWeekRow = hit.Row
End Function

Public Function isWeekClosed(ByVal ol As cOrderLine) As Boolean
    Dim ws As Worksheet, uge As Long, år As Long, ugeRow As Long
    
    uge = ol.uge
    år = ol.år
    
    Set ws = ThisWorkbook.Worksheets(LEVERINGSPLAN_PREFIX & år)
    ugeRow = FindWeekRow(ws, uge, år)
    
    isWeekClosed = LCase(ws.Cells(ugeRow, PLAN_COL_CLOSED)) = LCase("Lukket")
End Function

' Return True if the week is marked "Lukket" on the Leveringsplan sheet.
Public Function IsWeekClosedByYearWeek(ByVal yearNum As Long, ByVal weekNum As Long) As Boolean
    Dim ws As Worksheet, ugeRow As Long, cellVal As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEVERINGSPLAN_PREFIX & yearNum)
    On Error GoTo 0
    If ws Is Nothing Then
        IsWeekClosedByYearWeek = True   ' safest: missing sheet = closed
        Exit Function
    End If

    ugeRow = FindWeekRow(ws, weekNum, yearNum)
    If ugeRow <= 0 Then
        IsWeekClosedByYearWeek = True   ' no header found = closed
        Exit Function
    End If

    cellVal = LCase$(Trim$(CStr(ws.Cells(ugeRow, PLAN_COL_CLOSED).Value)))
    IsWeekClosedByYearWeek = (cellVal = LCase$("Lukket"))
End Function


Public Function isOneWeekClosed(ByVal pOrderLines As cOrderLines) As cOrderLine
    Dim ol As cOrderLine, line As cOrderLine
    
    Set line = New cOrderLine
    line.år = 0
    line.uge = 0
    line.kategori = ""
    line.antal = 0
    line.Dato = Empty
    line.OrderNo = ""
    
    
    Set isOneWeekClosed = line
    For Each ol In pOrderLines.Items
        If isWeekClosed(ol) Then
            Set isOneWeekClosed = ol
            Exit For               ' optional, since no further work needed
        End If
    Next
End Function



' Return the row of the requested category header **inside this week section**.
' weekRow = row returned by FindWeekRow
Public Function FindCategoryRowUnderWeek(ws As Worksheet, _
        ByVal weekRow As Long, ByVal yearNum As Long, _
        ByVal categoryName As String) As Long

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, PLAN_COL_HEADER_A).End(xlUp).Row

    For r = weekRow + 1 To lastRow
        Dim aVal As String: aVal = TrimEx(ws.Cells(r, PLAN_COL_HEADER_A).Value2)

        If Len(aVal) = 0 Then GoTo ContinueLoop        ' blank row
        If IsWeekHeaderValue(aVal, yearNum) Then Exit For          ' next week ? stop search

        If StrComp(aVal, categoryName, vbTextCompare) = 0 Then
            FindCategoryRowUnderWeek = r
            Exit Function
        End If
ContinueLoop:
    Next r
End Function


' Return the next header row (either next category or next week) **after** startRow.
' If none found, returns 0 (caller may treat as end-of-sheet).
Public Function FindNextHeaderRow(ws As Worksheet, _
        ByVal startRow As Long, ByVal yearNum As Long) As Long

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, PLAN_COL_HEADER_A).End(xlUp).Row

    For r = startRow + 1 To lastRow
        Dim aVal As String: aVal = TrimEx(ws.Cells(r, PLAN_COL_HEADER_A).Value2)
        If Len(aVal) = 0 Then GoTo ContinueLoop

        If IsWeekHeaderValue(aVal, yearNum) Or IsCategoryHeaderValue(aVal) Then
            FindNextHeaderRow = r
            Exit Function
        End If
ContinueLoop:
    Next r
End Function


' --------------- INTERNAL HELPERS ---------------

Public Function IsWeekHeaderValue(ByVal aVal As String, ByVal yearNum As Long) As Boolean
    Dim p As Long: p = InStr(aVal, "-")
    If p = 0 Then Exit Function

    Dim yTxt As String: yTxt = Mid$(aVal, p + 1)
    If Len(yTxt) = 4 And IsNumeric(yTxt) Then
        If CLng(yTxt) = yearNum Then IsWeekHeaderValue = True
    End If
End Function

Public Function IsCategoryHeaderValue(ByVal aVal As String) As Boolean
    Select Case True
        Case StrComp(aVal, CAT_PROD_MAL_SAME, vbTextCompare) = 0
        Case StrComp(aVal, CAT_PROD_SAME, vbTextCompare) = 0
        Case StrComp(aVal, CAT_PROD_THIS_NEXT, vbTextCompare) = 0
        Case StrComp(aVal, CAT_Q_STOP, vbTextCompare) = 0
        Case StrComp(aVal, CAT_LAGER, vbTextCompare) = 0
        Case StrComp(aVal, CAT_DIVERSE, vbTextCompare) = 0
        Case Else: Exit Function
    End Select
    IsCategoryHeaderValue = True
End Function

' Re-use the robust TrimEx from modMasterCache so we ignore NBSP, tabs, etc.
Private Function TrimEx(ByVal v As Variant) As String
    TrimEx = Replace$(Replace$(Replace$(Replace$(CStr(v), Chr$(160), " "), _
                      vbTab, " "), vbCr, " "), vbLf, " ")
    TrimEx = VBA.Trim$(TrimEx)
End Function

Public Function BuildWeeklyGaugeData(lines As cOrderLines) As Collection
    '–– 0) load just the weeks we need, pulling “used” and “total” straight from the plan
    modCapacity.LoadCapacityForLines lines

    Dim dictWeeks As Object
    Set dictWeeks = lines.GroupByWeekYear   ' keyed "YYYY|WW"
    
    Dim result    As New Collection
    Dim key       As Variant, parts As Variant
    Dim g         As cGauge
    Dim colLines  As Collection
    Dim ol        As cOrderLine
    Dim totalOrder As Double
    Dim yrNum     As Long, wkNum As Long
    
    For Each key In dictWeeks.Keys
        parts = Split(key, "|")
        yrNum = CLng(parts(0))
        wkNum = CLng(parts(1))
        
        ' sum only new orders
        totalOrder = 0
        Set colLines = dictWeeks(key)
        For Each ol In colLines
            totalOrder = totalOrder + ol.TotalHours
        Next

        Set g = New cGauge
        g.WeekKey = "Uge " & Format(wkNum, "0") & " – " & yrNum
        g.Capacity = modCapacity.GetWeekCapacityTotal(yrNum, wkNum)
        g.Used = modCapacity.GetWeekCapacityUsed(yrNum, wkNum)
        g.OrderLoad = totalOrder
        g.Remaining = g.Capacity - (g.Used + g.OrderLoad)
        g.IsOverflowing = (g.Used + g.OrderLoad > g.Capacity)
        
        result.Add g
    Next
    
    Set BuildWeeklyGaugeData = result
End Function

Public Sub GetIsoWeekYear(ByVal d As Date, ByRef isoWeek As Long, ByRef isoYear As Long)
    ' ISO 8601: weeks start Monday, week 1 contains Jan 4.
    Dim monday As Date:   monday = d - (Weekday(d, vbMonday) - 1)
    Dim thursday As Date: thursday = DateAdd("d", 3, monday)   ' decides ISO year
    isoYear = year(thursday)

    Dim jan4 As Date: jan4 = DateSerial(isoYear, 1, 4)
    Dim week1Mon As Date: week1Mon = jan4 - (Weekday(jan4, vbMonday) - 1)

    isoWeek = ((monday - week1Mon) \ 7) + 1
End Sub

