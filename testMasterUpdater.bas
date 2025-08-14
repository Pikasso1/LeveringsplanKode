Attribute VB_Name = "testMasterUpdater"
' ========== testMasterUpdater.bas ==========
Option Explicit

Public Sub testMasterUpdater()
    Dim wb          As Workbook
    Dim wsMaster    As Worksheet
    Dim dict        As Object         ' Scripting.Dictionary
    Dim duplicates  As Collection
    Dim lastRow     As Long
    Dim r           As Long
    Dim prodID      As Variant        ' ? must be Variant for For Each
    Dim msg         As String
    
    Set wb = ThisWorkbook
    Set wsMaster = wb.Worksheets(MASTER_SHEET_NAME)
    
    ' 1) Run the real update routine
    RefreshMasterProductList
    
    ' 2) Prepare for duplicate check
    Set dict = CreateObject("Scripting.Dictionary")
    Set duplicates = New Collection
    
    ' Find last used row via your Varenummer column
    lastRow = wsMaster.Cells(wsMaster.Rows.count, MASTER_COL_VARENR).End(xlUp).Row
    
    Debug.Print "Scanning Master sheet. MASTER_COL_VARENR=" & MASTER_COL_VARENR & _
            ", lastRow=" & lastRow
    
    ' 3) Scan for duplicates
    For r = MASTER_DATA_FIRST_ROW To lastRow
        prodID = CStr(wsMaster.Cells(r, MASTER_COL_VARENR).Value)
        If Len(prodID) > 0 Then
            If dict.Exists(prodID) Then
                On Error Resume Next
                  duplicates.Add prodID, prodID  ' avoid adding same dup twice
                On Error GoTo 0
            Else
                dict.Add prodID, r
            End If
        End If
    Next r
    
    ' 4) Report
    If duplicates.count = 0 Then
        Debug.Print "PASS: All product IDs are unique."
    Else
        msg = "FAIL: – duplicate IDs found:" & vbCrLf
        For Each prodID In duplicates   ' now compiles cleanly
            msg = msg & "  • " & prodID & " - " & vbCrLf
        Next
        Debug.Print msg
    End If
End Sub

