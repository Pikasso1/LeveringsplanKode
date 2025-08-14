Attribute VB_Name = "modFormShowerButton"
' ========= modFormShowerButton.bas =========
Option Explicit


Public Sub Show_New_Line_Form()
    clsLeveringsPlanForm.Show vbModeless
    Call RefreshMasterProductList
End Sub

