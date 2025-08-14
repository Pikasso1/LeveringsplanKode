Attribute VB_Name = "testPlanner"
' ========= TestPlanner.bas =========
Option Explicit

Public Sub Test_Plan_FindRows()
    Const yr As Long = 2025
    Const wk As Long = 34
    Const cat As String = CAT_LAGER   ' pick one of the 5 constants

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LEVERINGSPLAN_PREFIX & yr)

    Dim wr As Long: wr = FindWeekRow(ws, wk, yr)
    Dim cr As Long: cr = FindCategoryRowUnderWeek(ws, wr, yr, cat)
    Dim nextHdr As Long: nextHdr = FindNextHeaderRow(ws, cr, yr)

    Debug.Print "WeekRow", wr, "CatRow", cr, "NextHeader", nextHdr, "Category", cat
    If wr < cr And cr < nextHdr Then
        Debug.Print "PASS: Insert area found in row interval " & cr & "-" & nextHdr & " for Uge " & wk & "-" & yr
    Else
        Debug.Print "FAIL: Did not find suitable insert area for Uge " & wk & "-" & yr, "Category: " & cat
    End If
        
    If wr = 0 Then MsgBox "Week header not found"
    If cr = 0 Then MsgBox "Category header not found"
End Sub
