Attribute VB_Name = "modGuard"
Option Explicit

Public Function TryParseLongStrict(ByVal s As Variant, ByRef outVal As Long) As Boolean
    Dim t As String, i As Long
    t = Replace$(CStr(s), ChrW$(160), " ") ' remove NBSP
    t = Trim$(t)
    If Len(t) = 0 Then Exit Function
    
    ' allow optional leading minus sign
    For i = IIf(Left$(t, 1) = "-", 2, 1) To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then
            Exit Function
        End If
    Next i
    
    On Error GoTo EH
    outVal = CLng(t)
    TryParseLongStrict = True
    Exit Function
EH:
End Function

