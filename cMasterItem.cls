VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMasterItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ========= cMasterItem.cls =========
Option Explicit

Private mVarenr       As String
Private mNavn         As String
Private mHoursPerItem As Double
Private mPaintClass   As String
Private mProdNote     As String

Public Sub Init(ByVal varenr As String, ByVal Navn As String, _
                ByVal HoursPerItem As Double, ByVal paintClass As String, _
                ByVal ProdNote As String)
    mVarenr = varenr
    mNavn = Navn
    mHoursPerItem = HoursPerItem
    mPaintClass = paintClass
    mProdNote = ProdNote
End Sub

'-- NEW: internal constructor for placeholders
Friend Sub InitNew(vnr As String)
    mVarenr = vnr
    ' leave the rest blank/0
End Sub

Public Property Get varenr() As String
    varenr = mVarenr
End Property

Public Property Get Navn() As String
    Navn = mNavn
End Property

Public Property Get HoursPerItem() As Double
    HoursPerItem = mHoursPerItem
End Property

Public Property Get paintClass() As String
    paintClass = mPaintClass
End Property

Public Property Get ProdNote() As String
    ProdNote = mProdNote
End Property


