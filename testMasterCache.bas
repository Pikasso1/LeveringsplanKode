Attribute VB_Name = "testMasterCache"
' ========= TestMasterCache.bas =========
Option Explicit

Public Sub Test_MasterCache_Smoke()
    On Error GoTo FAIL

    InvalidateMasterCache
    InitMasterCache
    Debug.Assert MasterCount() >= 0

    ' Replace with a known varenr from your Master:
    Dim it As cMasterItem
    Set it = GetItem("9756998")
    Debug.Print "Test get item: ", it.Navn, it.HoursPerItem

    Dim tmp As cMasterItem, ok As Boolean
    ok = TryGetItem("___definitely_missing___", tmp)
    Debug.Assert ok = False
    Debug.Assert tmp Is Nothing

    Debug.Print "PASS: Master cache (class) - Items:", MasterCount()
    Exit Sub

FAIL:
    Debug.Print "FAIL: Master cache test: - " & Err.Description
    Err.Clear
End Sub

