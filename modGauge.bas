Attribute VB_Name = "modGauge"
' === modGauge.bas ===
Option Explicit

' Build weekly gauge data (wrapper around modPlanner)
Public Function GetGauges(lines As cOrderLines) As Collection
    Set GetGauges = modPlanner.BuildWeeklyGaugeData(lines)
End Function

' Clear only the dynamically added rows & labels …
Public Sub ClearContainer(fra As MSForms.Frame)
    Dim i   As Long
    Dim ctl As MSForms.Control

    ' 0-based collection ? indices run 0 To Count-1
    For i = fra.Controls.count - 1 To 0 Step -1
        Set ctl = fra.Controls.item(i)
        If Left$(ctl.name, 12) = "fraGaugeRow_" _
        Or Left$(ctl.name, 8) = "lblWeek_" Then
            fra.Controls.Remove i
        End If
    Next i
End Sub

' Render all gauges into a frame
Public Function RenderContainer( _
    fra As MSForms.Frame, _
    lines As cOrderLines, _
    parentForm As Object _
) As Long
    Dim gauges As Collection, g As cGauge
    Dim i       As Long

    ' 1) clear only inside the frame
    ClearContainer fra

    ' 2) add rows
    Set gauges = GetGauges(lines)
    For i = 1 To gauges.count
        Set g = gauges(i)
        AddRow fra, i, g
    Next i

    ' 3) size the frame & return its height
    fra.Height = gauges.count * CFG_ROW_HEIGHT
    RenderContainer = fra.Height
End Function

' Add a single row (frame + week label + 3 bars) into container `fra`
Private Sub AddRow(fra As MSForms.Frame, idx As Long, g As cGauge)
    Dim fraRow    As MSForms.Frame
    Dim lblW      As MSForms.Label, lbl As MSForms.Label
    Dim usableW   As Long, leftPos As Long
    Dim arr       As Variant, val As Double
    Dim barWidth  As Long
    Dim txt       As String
    Dim tip       As String

    ' create the row frame
    Set fraRow = fra.Controls.Add("Forms.Frame.1", "fraGaugeRow_" & idx)
    With fraRow
        .Top = (idx - 1) * CFG_ROW_HEIGHT
        .Left = 0
        .Width = fra.Width
        .Height = CFG_ROW_HEIGHT
    End With

    ' week label (nameplate)
    Set lblW = fraRow.Controls.Add("Forms.Label.1", "lblWeek_" & idx)
    With lblW
        .Caption = g.WeekKey
        .Left = 0
        .Width = CFG_WEEK_LABEL_WIDTH
        .Height = CFG_ROW_HEIGHT
        .TextAlign = fmTextAlignCenter

        ' --- CLOSED/WARNING COLORING ---
        Dim yr As Long, wk As Long, isClosed As Boolean
        If TryParseWeekKey(g.WeekKey, yr, wk) Then
            isClosed = modPlanner.IsWeekClosedByYearWeek(yr, wk)
        Else
            ' If parsing fails, be conservative
            isClosed = True
        End If

        If g.IsOverflowing Or isClosed Then
            .BackColor = CFG_COLOR_OVERFLOW_LABEL   ' red per config
            If isClosed Then
                tip = "CLOSED WEEK" & vbCrLf
            Else
                tip = ""
            End If
            tip = tip & "Capacity: " & g.Capacity & " hrs | Used: " & g.Used & _
                        " hrs | Orders: " & g.OrderLoad & " hrs | " & _
                        IIf(g.IsOverflowing, "Overflow: " & (g.Used + g.OrderLoad - g.Capacity) & " hrs", "No overflow")
        Else
            .BackColor = CFG_COLOR_DEFAULT_LABEL
            tip = "Capacity: " & g.Capacity & " hrs | Used: " & g.Used & _
                  " hrs | Orders: " & g.OrderLoad & " hrs"
        End If

        .ControlTipText = tip
    End With

    ' gauge bars (green Used, red Order, yellow Remaining)
    usableW = fraRow.Width - CFG_WEEK_LABEL_WIDTH
    leftPos = CFG_WEEK_LABEL_WIDTH

    ' OPTIONAL: show bars even if capacity = 0 (closed weeks).
    ' Leave False to keep your current behavior (no bars when capacity=0).
    Const SHOW_BARS_WHEN_CLOSED As Boolean = False

    For Each arr In Array( _
        Array("Used", g.Used), _
        Array("Order", g.OrderLoad), _
        Array("Remain", g.Remaining) _
      )
        val = arr(1)
        txt = Format(val, "#,##0")

        Set lbl = fraRow.Controls.Add("Forms.Label.1", "lbl" & arr(0) & "_" & idx)
        With lbl
            .Left = leftPos
            .Height = CFG_ROW_HEIGHT

            ' compute + clamp width
            Dim denom As Double
            If g.Capacity > 0 Then
                denom = g.Capacity
            ElseIf SHOW_BARS_WHEN_CLOSED Then
                ' fallback scale by activity so something renders
                denom = g.Used + g.OrderLoad
                If denom <= 0 Then denom = 1 ' avoid /0
            Else
                denom = 0
            End If

            If denom > 0 Then
                barWidth = usableW * (val / denom)
            Else
                barWidth = 0
            End If

            If barWidth < 0 Then barWidth = 0
            If barWidth > usableW Then barWidth = usableW
            .Width = barWidth

            ' show the value centered
            .Caption = txt
            .TextAlign = fmTextAlignCenter

            ' fill colors from config
            Select Case arr(0)
                Case "Used":   .BackColor = CFG_COLOR_USED_FILL
                Case "Order":  .BackColor = CFG_COLOR_ORDER_FILL
                Case "Remain": .BackColor = CFG_COLOR_REMAIN_FILL
            End Select
        End With

        leftPos = leftPos + barWidth
    Next
End Sub

' Return (weekKey, overByHours) pairs for any overflow
Public Function GetOverflowTuples(lines As cOrderLines) As Collection
    Dim res As New Collection
    Dim gauges As Collection, g As cGauge
    Dim tup As Variant, overBy As Double

    Set gauges = GetGauges(lines)

    For Each g In gauges
        If g.IsOverflowing Then
            overBy = (g.Used + g.OrderLoad - g.Capacity)
            If overBy > 0 Then
                tup = Array(g.WeekKey, CDbl(overBy))
                res.Add tup
            End If
        End If
    Next

    Set GetOverflowTuples = res
End Function

' Parse "Uge 34 – 2025" -> wk=34, yr=2025 (robust to NBSP + fancy dashes)
Private Function TryParseWeekKey(ByVal key As String, _
                                 ByRef yr As Long, _
                                 ByRef wk As Long) As Boolean
    Dim t As String, dashPos As Long, wTxt As String, yTxt As String
    t = LCase$(Trim$(Replace$(key, ChrW$(160), "")))
    t = Replace$(t, "–", "-")        ' en dash -> hyphen
    t = Replace$(t, "—", "-")        ' em dash -> hyphen
    t = Replace$(t, "uge", "")
    t = Replace$(t, " ", "")

    dashPos = InStr(1, t, "-")
    If dashPos = 0 Then Exit Function

    wTxt = Mid$(t, 1, dashPos - 1)
    yTxt = Mid$(t, dashPos + 1)

    If TryParseLongStrict(wTxt, wk) And TryParseLongStrict(yTxt, yr) Then
        If wk >= 1 And wk <= 52 And yr >= 2020 And yr <= 2100 Then
            TryParseWeekKey = True
        End If
    End If
End Function


