Attribute VB_Name = "modMasterCache"
' ========= modMasterCache.bas =========
Option Explicit

Private pCache As Object  ' Scripting.Dictionary: key As String -> cMasterItem

Public Sub InitMasterCache()
    Dim ws As Worksheet
    Set ws = SheetByNameSafe(MASTER_SHEET_NAME)
    If ws Is Nothing Then Err.Raise vbObjectError + 6101, , _
        "Master sheet '" & MASTER_SHEET_NAME & "' not found."

    Dim lastRow As Long
    lastRow = LastUsedRowIn(ws, MASTER_COL_VARENR)

    Set pCache = CreateObject("Scripting.Dictionary")
    pCache.CompareMode = 1 ' TextCompare

    If lastRow < MASTER_DATA_FIRST_ROW Then Exit Sub

    Dim r As Long
    For r = MASTER_DATA_FIRST_ROW To lastRow
        Dim rawKey As Variant
        rawKey = ws.Cells(r, MASTER_COL_VARENR).Value2
        If Not IsEmpty(rawKey) And Len(CStr(rawKey)) > 0 Then
            Dim key As String: key = NormalizeKey(CStr(rawKey))

            If Not pCache.Exists(key) Then
                Dim itm As cMasterItem
                Set itm = New cMasterItem
                itm.Init key, _
                         NzString(ws.Cells(r, MASTER_COL_NAVN).Value2), _
                         NzDouble(ws.Cells(r, MASTER_COL_HOURS).Value2), _
                         NzString(ws.Cells(r, MASTER_COL_PAINT).Value2), _
                         NzString(ws.Cells(r, MASTER_COL_PRODNOTE).Value2)
                pCache.Add key, itm
            ElseIf Not MASTER_ALLOW_DUP_VNR Then
                ' First seen wins; optionally Debug.Print the duplicate
            End If
        End If
    Next r
End Sub

Public Sub InvalidateMasterCache()
    Set pCache = Nothing
End Sub

' Try-get (no error)
Public Function TryGetItem(ByVal varenr As String, ByRef outItem As cMasterItem) As Boolean
    EnsureCache
    Dim key As String: key = NormalizeKey(varenr)
    If pCache.Exists(key) Then
        Set outItem = pCache(key)
        TryGetItem = True
    Else
        Set outItem = Nothing
        TryGetItem = False
    End If
End Function

' GetItem (throws if missing)
Public Function GetItem(ByVal varenr As String) As cMasterItem
    EnsureCache
    Dim key As String: key = NormalizeKey(varenr)
    If Not pCache.Exists(key) Then
        ' Err.Raise vbObjectError + 6102, , "Varenr not in Master: '" & varenr & "' (normalized: '" & key & "')"
    End If
    Set GetItem = pCache(key)
End Function

Public Function MasterCount() As Long
    EnsureCache
    MasterCount = pCache.count
End Function

' --- helpers ---
Private Sub EnsureCache()
    If pCache Is Nothing Then InitMasterCache
End Sub

Private Function SheetByNameSafe(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetByNameSafe = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
End Function

Private Function LastUsedRowIn(ws As Worksheet, ByVal col As Long) As Long
    LastUsedRowIn = ws.Cells(ws.Rows.count, col).End(xlUp).Row
End Function

' --- helpers: robust whitespace handling ---
Private Function TrimEx(ByVal s As String) As String
    ' Normalize NBSP to space, strip tabs & control chars, then Trim$
    s = Replace$(s, Chr$(160), " ")     ' NBSP -> space
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    TrimEx = Trim$(s)
End Function

Private Function IsBlankV(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        IsBlankV = True
    Else
        IsBlankV = (Len(TrimEx(CStr(v))) = 0)
    End If
End Function

Private Function NormalizeKey(ByVal raw As String) As String
    Dim s As String
    s = TrimEx(raw)
    If MASTER_UPPERCASE_KEYS Then s = UCase$(s)
    NormalizeKey = s
End Function

Private Function NzString(ByVal v As Variant) As String
    If IsBlankV(v) Then
        NzString = ""
    Else
        NzString = TrimEx(CStr(v))
    End If
End Function

Private Function NzDouble(ByVal v As Variant) As Double
    If IsBlankV(v) Then
        NzDouble = 0#
    Else
        NzDouble = CDbl(v)
    End If
End Function

' Return a blank Master item when Varenr not found
Public Function CreatePlaceholderItem(varenr As Variant) As cMasterItem
    Dim tmp As New cMasterItem
    tmp.InitNew CStr(varenr)             ' only Friend code may call this
    Set CreatePlaceholderItem = tmp
End Function





