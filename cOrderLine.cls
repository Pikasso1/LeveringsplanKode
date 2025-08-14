VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOrderLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ========= cOrderLine.cls =========
Option Explicit

Private mMaster As cMasterItem

' Staging fields the user controls
Public år               As Long
Public uge              As Long
Public kategori         As String
Public antal            As Double
Public Dato             As Variant   ' Empty or Date
Public OrderNo          As String    ' optional, for writer later
Public SourceRow        As Long      ' optional, for notPlanned later
Public SourceSheet      As String    ' optional, for notPlanned later
Public paintCatNeeded   As Boolean   ' optional, for category/commit guarding later


' --- init / attach master record ---
Public Sub InitFromMaster(ByVal masterItem As cMasterItem)
    Set mMaster = masterItem
End Sub

' --- read-only views into master data ---
Public Property Get varenr() As String
    varenr = mMaster.varenr
End Property

Public Property Get Navn() As String
    Navn = mMaster.Navn
End Property

Public Property Get HoursPerItem() As Double
    HoursPerItem = mMaster.HoursPerItem
End Property

Public Property Get paintClass() As String
    paintClass = mMaster.paintClass
End Property

Public Property Get ProdNote() As String
    ProdNote = mMaster.ProdNote
End Property

Public Property Get TotalHours() As Double
    TotalHours = antal * mMaster.HoursPerItem + 1.5
End Property


' --- per-line validation ---
Public Function Validate(ByRef outIssue As String) As Boolean
    If år = 0 Or uge = 0 Or Len(Trim$(kategori)) = 0 Then
        outIssue = "Missing År/Uge/Kategori."
        Exit Function
    End If
    If RequiresDato(kategori) Then
        If IsEmpty(Dato) Or IsNull(Dato) Or Len(Trim$(CStr(Dato))) = 0 Then
            outIssue = "Dato required for '" & CAT_PROD_THIS_NEXT & "'."
            Exit Function
        End If
    End If
    If mMaster Is Nothing Then
        outIssue = "Missing Master data."
        Exit Function
    End If
    If antal <= 0 Then
        outIssue = "Antal must be > 0."
        Exit Function
    End If
    
    ' Hours per item validation fails for new cylinders, thus its commented out as a TODO fix
    'If mMaster.HoursPerItem <= 0 Then
    '    outIssue = "HoursPerItem missing/zero in Master for Varenr=" & mMaster.varenr
    '    Exit Function
    'End If

    Validate = True
End Function

