VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTrajectoryCrossingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("class to encapsulate trajectory crossing details")
'
'                (x2,y2)
'                 |..................
'                 |
'   ---------+----+----------------------------------  Test Elevation
'            |
'   .........|
'          (x1,y1)
'
'   ---------------------------------------------------- Line of Sight

Private m_x1 As Double
Private m_y1 As Double
Private m_x2 As Double
Private m_y2 As Double
Private m_index1 As Long
Private m_index2 As Long


Public Function Initialize(detail As TCrossingDetail) As CTrajectoryCrossingDetail
    Me.SetDetail detail
    Set Initialize = Me
End Function


Public Sub SetDetail(detail As TCrossingDetail)
    With detail
        m_x1 = .pX1
        m_y1 = .pY1
        m_index1 = .pIndex1
        
        m_x2 = .pX2
        m_y2 = .pY2
        m_index2 = .pIndex2
    End With
End Sub

Public Property Get X1() As Double
    X1 = m_x1
End Property

Public Property Get Y1() As Double
    Y1 = m_y1
End Property

Public Property Get Index1() As Long
    Index1 = m_index1
End Property

Public Property Get X2() As Double
    X2 = m_x2
End Property

Public Property Get Y2() As Double
    Y2 = m_y2
End Property

Public Property Get Index2() As Long
    Index2 = m_index2
End Property

'@Description("find the range at which the trajectory actually intersects the target elevation line")
'
Public Function RangeAtTargetY(ByVal targetY As Double, ByRef targetX As Double) As Boolean

    ' check to see that the specified targetY value is bracketed by the crossing points
    If Not Me.ContainsTargetY(targetY) Then
        targetX = 0#
        RangeAtTargetY = False
        Exit Function
    End If

    ' just use straightforward linear interpolation
    ' y = a0 + a1 * x
    Dim a1 As Double: a1 = (m_y2 - m_y1) / (m_x2 - m_x1)
    Dim a0 As Double: a0 = (m_x2 * m_y1 - m_x1 * m_y2) / (m_x2 - m_x1)
    
    targetX = (targetY - a0) / a1
    
    RangeAtTargetY = True
    
End Function

Public Function ContainsTargetY(ByVal targetY As Double) As Boolean

    Dim result As Boolean
    
    If targetY >= m_y1 And targetY <= m_y2 Then
        result = True
    ElseIf targetY >= m_y2 And targetY <= m_y1 Then
        result = True
    Else
        result = False
    End If
    
    ContainsTargetY = result

End Function
