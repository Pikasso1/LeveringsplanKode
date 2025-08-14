VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOrderLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ========= cOrderLines.cls =========
Option Explicit

Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

'— count —
Public Property Get count() As Long
    count = mItems.count
End Property

'— indexer —
Public Property Get item(ByVal index As Long) As cOrderLine
    Set item = mItems(index)
End Property

'— expose underlying collection for For Each —
Public Property Get Items() As Collection
    Set Items = mItems
End Property

'— add & enrich —
Public Function AddLine(ByVal varenr As String, _
                        ByVal antal As Double, _
                        Optional ByVal SourceRow As Long = 0, _
                        Optional ByVal SourceSheet As String = "", _
                        Optional ByVal paintCatNeeded As Boolean) As cOrderLine

    Dim mi As cMasterItem
    Set mi = modMasterCache.GetItem(varenr)

    Dim ol As New cOrderLine
    ol.InitFromMaster mi
    ol.antal = antal
    ol.SourceRow = SourceRow            ' optional
    ol.SourceSheet = SourceSheet        ' optional
    ol.paintCatNeeded = paintCatNeeded  ' optional

    mItems.Add ol
    Set AddLine = ol
End Function

'— weekly sum —
Public Function SumHours(ByVal yearNum As Long, ByVal weekNum As Long) As Double
    Dim ol As cOrderLine
    For Each ol In mItems
        If ol.år = yearNum And ol.uge = weekNum Then _
            SumHours = SumHours + ol.TotalHours
    Next
End Function

'— validation —
Public Function ValidateAll(ByRef issues As Collection) As Boolean
    Dim ok As Boolean: ok = True
    Set issues = New Collection

    Dim i As Long, ol As cOrderLine, msg As String
    For Each ol In mItems
        i = i + 1
        If Not ol.Validate(msg) Then
            ok = False
            issues.Add "Line " & i & ": " & msg
        End If
    Next
    ValidateAll = ok
End Function

' Return Dictionary keyed "yyyy|ww", value = Collection of cOrderLine
Public Function GroupByWeekYear() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ol As cOrderLine, k$
    For Each ol In mItems
        k = ol.år & "|" & ol.uge
        If Not d.Exists(k) Then Set d(k) = New Collection
        d(k).Add ol
    Next
    Set GroupByWeekYear = d
End Function

' Create from an already-loaded Master item (no cache lookup)
Public Function AddLineFromMaster(mi As cMasterItem, qty As Double) As cOrderLine
    Dim ol As New cOrderLine
    ol.InitFromMaster mi
    ol.antal = qty
    mItems.Add ol
    Set AddLineFromMaster = ol
End Function




