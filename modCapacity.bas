Attribute VB_Name = "modCapacity"
' ========= modCapacity.bas =========
Option Explicit

'— caches for only the weeks we care about
Private pCapUsed  As Object
Private pCapTotal As Object

' --- single source of truth for dictionary keys ---
Private Function CapKey(ByVal yr As Long, ByVal wk As Long) As String
    ' Use non-padded week consistently, e.g., "2025|2".
    ' (If you prefer zero-padded, switch to Format$(wk, "00") here AND nowhere else.)
    CapKey = CStr(yr) & "|" & CStr(wk)
End Function

' Call this once with your staged OrderLines to prime the caches
Public Sub LoadCapacityForLines(lines As cOrderLines)
    Dim dictWeeks As Object
    Set dictWeeks = lines.GroupByWeekYear    ' keys = "YYYY|WW" (WW might be padded)

    ' initialize
    Set pCapUsed = CreateObject("Scripting.Dictionary")
    Set pCapTotal = CreateObject("Scripting.Dictionary")

    Dim key As Variant, parts As Variant
    Dim yr As Long, wk As Long
    Dim ws As Worksheet
    Dim weekRow As Long
    Dim k As String

    For Each key In dictWeeks.Keys
        parts = Split(key, "|")
        yr = CLng(parts(0))
        wk = CLng(parts(1))            ' <- CLng removes any zero padding
        k = CapKey(yr, wk)             ' <- normalize

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(LEVERINGSPLAN_PREFIX & yr)
        On Error GoTo 0

        If ws Is Nothing Then
            pCapUsed(k) = 0
            pCapTotal(k) = 0
        Else
            weekRow = modPlanner.FindWeekRow(ws, wk, yr)
            If weekRow > 0 Then
                pCapUsed(k) = NzNum(ws.Cells(weekRow, PLAN_COL_CAPACITY_USED).Value)
                pCapTotal(k) = NzNum(ws.Cells(weekRow, PLAN_COL_CAPACITY_TOTAL).Value)
            Else
                pCapUsed(k) = 0
                pCapTotal(k) = 0
            End If
        End If

        Set ws = Nothing
    Next
End Sub

' Returns “already used” for that week (must have called LoadCapacityForLines first)
Public Function GetWeekCapacityUsed(ByVal yr As Long, ByVal wk As Long) As Double
    Dim k As String: k = CapKey(yr, wk)
    If pCapUsed Is Nothing Then Err.Raise 5, , "Call LoadCapacityForLines first"
    If pCapUsed.Exists(k) Then GetWeekCapacityUsed = pCapUsed(k)
End Function

' Returns total capacity for that week
Public Function GetWeekCapacityTotal(ByVal yr As Long, ByVal wk As Long) As Double
    Dim k As String: k = CapKey(yr, wk)
    If pCapTotal Is Nothing Then Err.Raise 5, , "Call LoadCapacityForLines first"
    If pCapTotal.Exists(k) Then GetWeekCapacityTotal = pCapTotal(k)
End Function

' helper: treat non-numeric as zero
Private Function NzNum(v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v) Else NzNum = 0
End Function

' --- optional debug helper: call from Immediate window if needed ---
Public Sub DebugDumpCapacityCaches()
    Dim k As Variant
    Debug.Print "---- Capacity cache dump ----"
    If Not pCapTotal Is Nothing Then
        For Each k In pCapTotal.Keys
            Debug.Print k, "Total=", pCapTotal(k), " Used=", IIf(pCapUsed.Exists(k), pCapUsed(k), "N/A")
        Next k
    Else
        Debug.Print "(pCapTotal is Nothing)"
    End If
End Sub

