Attribute VB_Name = "modMaling"
' === modMaling.bas ===
Option Explicit

Public Function RequiresDato(ByVal category As String) As Boolean
    RequiresDato = _
        (StrComp(category, CAT_PROD_THIS_NEXT, vbTextCompare) = 0) _
     Or (StrComp(category, CAT_Q_STOP, vbTextCompare) = 0)
End Function

Public Function CategoryNeedsPaint(ByRef ol As cOrderLine) As Boolean
    If ol.paintClass <> "" Then
        ' Has to be in red, green, pink or blue
        If ol.kategori <> CAT_PROD_MAL_SAME _
        And ol.kategori <> CAT_PROD_THIS_NEXT _
        And ol.kategori <> CAT_Q_STOP _
        And ol.kategori <> CAT_LAGER Then
            ol.paintCatNeeded = True
            CategoryNeedsPaint = True
        End If
    Else
        CategoryNeedsPaint = False
    End If
End Function

Public Function CheckCategoryNeedsPaint(ByVal pOrderLines As cOrderLines) As cOrderLine
    Dim ol As cOrderLine, line As cOrderLine
    
    Set line = New cOrderLine
    line.år = 0
    line.uge = 0
    line.kategori = ""
    line.antal = 0
    line.Dato = Empty
    line.OrderNo = ""
    
    Set CheckCategoryNeedsPaint = line
    For Each ol In pOrderLines.Items
        If CategoryNeedsPaint(ol) Then
            Set CheckCategoryNeedsPaint = ol
            Exit For
        End If
    Next
    
End Function
