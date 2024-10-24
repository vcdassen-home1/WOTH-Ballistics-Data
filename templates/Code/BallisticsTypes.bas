Attribute VB_Name = "BallisticsTypes"
Option Explicit

Public Type TCrossingDetail
    pX1 As Double
    pY1 As Double
    pIndex1 As Long
    
    pX2 As Double
    pY2 As Double
    pIndex2 As Double
    
End Type

Public Enum ETargetDiameterCrossing

    Indeterminate = -2
    BelowTarget = -1
    InsideTarget = 0
    AboveTarget = 1
    
    
End Enum

Public Enum ELOSCrossing

    Indeterminate = -2
    BelowLOS = -1
    AtLOS = 0
    AboveLOS = 1
    
    
End Enum

Public Function TrajectoryTracker(ByVal nPoints As Long, ByVal yMin As Double, ByVal yMax As Double) As CTrajectoryTracker
    Dim obj As New CTrajectoryTracker
    Set TrajectoryTracker = obj.Initialize(nPoints, yMin, yMax)
End Function

Public Function TrajectoryCrossingDetail(detail As TCrossingDetail) As CTrajectoryCrossingDetail
    Dim obj As New CTrajectoryCrossingDetail
    Set TrajectoryCrossingDetail = obj.Initialize(detail)
End Function
