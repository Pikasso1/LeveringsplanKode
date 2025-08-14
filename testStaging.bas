Attribute VB_Name = "testStaging"
' ========= TestStaging.bas =========
Option Explicit

Public Sub Test_Staging_Smoke()
    InvalidateMasterCache
    InitMasterCache

    Dim lines As New cOrderLines
    Dim l1 As cOrderLine, l2 As cOrderLine

    ' Use real varenr from your Master
    Set l1 = lines.AddLine("9756998", 10)   ' <-- replace with real
    l1.år = 2025: l1.uge = 34: l1.kategori = CAT_PROD_SAME

    Set l2 = lines.AddLine("9752349", 5)    ' <-- replace with real
    l2.år = 2025: l2.uge = 34: l2.kategori = CAT_PROD_THIS_NEXT
    l2.Dato = DateSerial(2025, 8, 25)       ' required for this category

    Dim issues As Collection
    If Not lines.ValidateAll(issues) Then
        Dim it As Variant
        For Each it In issues: Debug.Print it: Next
        Err.Raise vbObjectError + 6201, , "Validation failed."
    End If

    Debug.Print "SumHours(2025,34) = "; lines.SumHours(2025, 34)
    Debug.Print "PASS: Staging smoke test:, lines=" & lines.count
End Sub

