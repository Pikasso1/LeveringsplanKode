Attribute VB_Name = "testDeliveryWriter"
' ========= testDeliveryWriter.bas =========
Option Explicit

Sub Test_Writer_Into_Plan()
InvalidateMasterCache:     InitMasterCache

    '--Build dummy staging lines--
    Dim lines As New cOrderLines
    Dim l As cOrderLine

    Set l = lines.AddLine("9756998", 3)   'adjust VNRs to real ones
    l.år = 2025: l.uge = 34: l.kategori = CAT_PROD_SAME
    l.Dato = Empty
    l.OrderNo = "029999"

    Set l = lines.AddLine("9752349", 4)
    l.år = 2025: l.uge = 35: l.kategori = CAT_PROD_SAME
    l.OrderNo = "029999"

    CommitStagedLines lines, "029999"
End Sub

