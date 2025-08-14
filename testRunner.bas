Attribute VB_Name = "testRunner"
' ========= testRunner.bas =========
Option Explicit

'--- Hard-coded list -------------------------------------------------
Private Sub RunAllTests()
    On Error GoTo BailOut
    ' Debug.Print " " are used as test seperators
    Debug.Print " "
    Debug.Print "============================================================ MASTER TEST RUN START ============================================================"
    
    ' Test entry points
    
    testMasterCache.Test_MasterCache_Smoke
    Debug.Print " "
    testMasterUpdater.testMasterUpdater
    Debug.Print " "
    testPlanner.Test_Plan_FindRows
    Debug.Print " "
    ' TODO Recomfigure writer test to write to every category in every week in every year, then delete them again
    ' testDeliveryWriter.Test_Writer_Into_Plan
    testStaging.Test_Staging_Smoke
    
    Debug.Print "============================================================ ALL TESTS COMPLETED ============================================================"
    MsgBox "All tests ran – see Immediate window for PASS / FAIL.", vbInformation
    Exit Sub
BailOut:
    MsgBox "Master test run aborted: " & Err.Description, vbCritical
End Sub

