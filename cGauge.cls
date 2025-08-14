VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === cGauge.cls ===
Option Explicit

'–– Private storage
Private pWeekKey    As String
Private pCapacity   As Double
Private pUsed       As Double
Private pOrderLoad  As Double
Private pRemaining  As Double
Private pIsOverflow As Boolean

'–– Properties
Public Property Get WeekKey() As String
    WeekKey = pWeekKey
End Property
Public Property Let WeekKey(val As String)
    pWeekKey = val
End Property

Public Property Get Capacity() As Double
    Capacity = pCapacity
End Property
Public Property Let Capacity(val As Double)
    pCapacity = val
End Property

Public Property Get Used() As Double
    Used = pUsed
End Property
Public Property Let Used(val As Double)
    pUsed = val
End Property

Public Property Get OrderLoad() As Double
    OrderLoad = pOrderLoad
End Property
Public Property Let OrderLoad(val As Double)
    pOrderLoad = val
End Property

Public Property Get Remaining() As Double
    Remaining = pRemaining
End Property
Public Property Let Remaining(val As Double)
    pRemaining = val
End Property

Public Property Get IsOverflowing() As Boolean
    IsOverflowing = pIsOverflow
End Property
Public Property Let IsOverflowing(val As Boolean)
    pIsOverflow = val
End Property

