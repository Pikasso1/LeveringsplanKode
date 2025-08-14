Attribute VB_Name = "modMasterUpdater"
' ========= modMasterUpdater.bas =========
Option Explicit

Private dict As Object            ' Scripting.Dictionary

Public Sub ButtonRefreshMasterList()
    Call RefreshMasterProductList
    
    MsgBox dict.count & " produkter opdateret i """ & MASTER_SHEET_NAME & """-arket.", vbInformation
End Sub

Public Sub RefreshMasterProductList()
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim wsMaster As Worksheet
    Dim idxList As Collection
    Dim i As Long, r As Long
    Dim lastRow As Long, lastCol As Long
    Dim prodID As String
    Dim rowData As Variant

    Set wb = ThisWorkbook
    Set wsMaster = wb.Worksheets(MASTER_SHEET_NAME)
    Set dict = CreateObject("Scripting.Dictionary")
    Set idxList = New Collection

    ' 1) Find all "Leveringsplan " sheets, newest first
    For i = wb.Worksheets.count To 1 Step -1
        Set sh = wb.Worksheets(i)
        If Left(sh.name, Len(LEVERINGSPLAN_PREFIX)) = LEVERINGSPLAN_PREFIX Then
            idxList.Add sh
        End If
    Next i

    ' 2) Scan each sheet in that order
    For Each sh In idxList
        ' Determine data bounds
        lastRow = sh.Cells(sh.Rows.count, PLAN_COL_VARENR).End(xlUp).Row
        lastCol = PLAN_COL_PRODNOTE

        For r = lastRow To 2 Step -1
            prodID = CStr(sh.Cells(r, PLAN_COL_VARENR).Value)
            
            ' Sometimes random shit gets put into my schema, idk man
            ' Cant really be bothered to guard it, since theres no good way to get everything then
            'If Not IsNumeric(prodID) Then
            '    Debug.Print "Shit broke at " & prodID
            '    prodID = ""
            'End If
            
            If MASTER_TRIM_KEYS Then prodID = Trim(prodID)
            If MASTER_UPPERCASE_KEYS Then prodID = UCase(prodID)

            If Len(prodID) > 0 Then
                If Not dict.Exists(prodID) Then
                    ' Cache the entire row into a 1×lastCol array
                    rowData = sh.Range( _
                        sh.Cells(r, 1), _
                        sh.Cells(r, lastCol) _
                    ).Value
                    dict.Add prodID, rowData
                ElseIf MASTER_ALLOW_DUP_VNR Then
                    ' If you want to allow duplicates, you could log or overwrite here
                    ' dict(prodID) = rowData
                End If
            End If
        Next r
    Next sh

    ' 3) Write results to Master sheet
    wsMaster.Cells.ClearContents

    ' 3a) Copy header row from the first Leveringsplan sheet
    If idxList.count > 0 Then
        idxList(1).Rows(1).Copy Destination:=wsMaster.Rows(1)
    End If

    ' 3b) Build output array
    Dim outArr() As Variant
    Dim key As Variant
    ReDim outArr(1 To dict.count, 1 To lastCol)

    i = 0
    For Each key In dict.Keys
        i = i + 1
        ' dict(key) is a 2D array (1 to 1, 1 to lastCol)
        For r = 1 To lastCol
            outArr(i, r) = dict(key)(1, r)
        Next r
    Next key

    ' 3c) Dump into Master sheet starting at MASTER_DATA_FIRST_ROW
    With wsMaster
        .Range( _
          .Cells(MASTER_DATA_FIRST_ROW, 1), _
          .Cells(MASTER_DATA_FIRST_ROW + dict.count - 1, lastCol) _
        ).Value = outArr
    End With
End Sub


