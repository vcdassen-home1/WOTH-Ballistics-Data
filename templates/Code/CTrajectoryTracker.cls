VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTrajectoryTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_trajectory() As Double
Private m_targetCrossings() As ETargetDiameterCrossing


Private m_nPoints As Long
Private m_yMin As Double
Private m_yMax As Double
Private m_index As Long


Private Sub Class_Initialize()
    ReDim m_trajectory(1 To 1, 1 To 1) As Double
    ReDim m_targetCrossings(1 To 1) As ETargetDiameterCrossing
    m_index = 1
End Sub

Private Sub Class_Terminate()

    Erase m_trajectory
    Erase m_targetCrossings

End Sub

Public Property Get Size() As Long
    Size = m_nPoints
End Property

Public Property Get yMax() As Double
    yMax = m_yMax
End Property

Public Property Get yMin() As Double
    yMin = m_yMin
    
End Property

Public Property Get trajectoryData() As Variant
    trajectoryData = m_trajectory
End Property

Public Property Get TargetCrossings() As Variant
    TargetCrossings = m_targetCrossings
End Property

Public Function Initialize(ByVal nPoints As Long, ByVal yMin As Double, ByVal yMax As Double) As CTrajectoryTracker

    m_yMin = yMin
    m_yMax = yMax
    
    m_nPoints = nPoints
    ReDim m_trajectory(1 To nPoints, 1 To 3) As Double
    ReDim m_targetCrossings(1 To nPoints) As ETargetDiameterCrossing
    InitializeArrays
    
    Set Initialize = Me
    
End Function

Public Sub Reset()
    m_index = 1
    InitializeArrays
End Sub

Private Sub InitializeArrays()

    Dim i As Long
    For i = 1 To m_nPoints
        m_trajectory(i, 1) = -1#
        m_trajectory(i, 2) = 0#
        m_trajectory(i, 3) = 0#
        
        m_targetCrossings(i) = ETargetDiameterCrossing.Indeterminate
    Next i

End Sub

Public Sub Track(ByVal x As Double, ByVal y As Double)

    If m_index >= 1 And m_index <= m_nPoints Then
    
        m_trajectory(m_index, 1) = x
        m_trajectory(m_index, 2) = y
        m_targetCrossings(m_index) = YCrossing(y)
        m_index = m_index + 1
        
    End If

End Sub

Public Sub TrackTrajectoryAndEnergy(ByVal x As Double, ByVal y As Double, ByVal energy As Double)

    If m_index >= 1 And m_index <= m_nPoints Then
    
        m_trajectory(m_index, 1) = x
        m_trajectory(m_index, 2) = y
        m_trajectory(m_index, 3) = energy
        m_targetCrossings(m_index) = YCrossing(y)
        m_index = m_index + 1
        
    End If

End Sub

Private Function YCrossing(ByVal y As Double) As ETargetDiameterCrossing

    Dim yCross As ETargetDiameterCrossing
    
    If y > m_yMax Then
        yCross = ETargetDiameterCrossing.AboveTarget
    ElseIf y < m_yMin Then
        yCross = ETargetDiameterCrossing.BelowTarget
    Else
        yCross = ETargetDiameterCrossing.InsideTarget
    End If
    
    YCrossing = yCross

End Function

Public Sub FindCrossings(crossingList As ArrayList)

    Dim i1 As Long
    Dim i2 As Long
    Dim i As Long
    Dim startIndex As Long: startIndex = LBound(m_trajectory, 1)
    Dim lastIndex As Long: lastIndex = UBound(m_trajectory, 1) - 1
    
    For i = startIndex To lastIndex
        i1 = i
        i2 = i + 1
        If Not m_targetCrossings(i2) = m_targetCrossings(i1) Then
            'found an index pair bounding a target crossing
            AddCrossingDetail i1, i2, crossingList
        End If
    Next i
    

End Sub

Private Sub AddCrossingDetail(ByVal i1 As Long, ByVal i2 As Long, crossingList As ArrayList)

    
    Dim detail As TCrossingDetail
    With detail
        .pIndex1 = i1
        .pIndex2 = i2
        .pX1 = m_trajectory(i1, 1)
        .pY1 = m_trajectory(i1, 2)
        .pX2 = m_trajectory(i2, 1)
        .pY2 = m_trajectory(i2, 2)
    End With
    crossingList.Add TrajectoryCrossingDetail(detail)
End Sub

'
'  For ballistic trajectories, there are only 5 possibilities
'  0 crossings
'  1 crossing
'  2 crossings
'  3 crossings
'  4 crossings

Public Sub BuildPointBlankRangeList(crossingList As ArrayList, pbrList As ArrayList)

    Select Case crossingList.Count
    
        Case 0
            Exit Sub
            
        Case 1
            ConstructPBRInterval1 crossingList, pbrList
            
        Case 2
        
            ConstructPBRInterval2 crossingList, pbrList
        
        Case 3
        
            ConstructPBRInterval3 crossingList, pbrList
        
        Case 4
        
            ConstructPBRInterval4 crossingList, pbrList
            
        Case Else
    
    End Select

End Sub


Private Sub ConstructPBRInterval1(crossingList As ArrayList, pbrList As ArrayList)

    Dim cross1 As CTrajectoryCrossingDetail: Set cross1 = crossingList(0)
    Dim pbr1 As CPBRData: Set pbr1 = ConstructPBRDataFromZeroToFirstCrossing(cross1)
    pbrList.Add pbr1

End Sub

Private Sub ConstructPBRInterval2(crossingList As ArrayList, pbrList As ArrayList)

    
    Dim cross1 As CTrajectoryCrossingDetail, cross2 As CTrajectoryCrossingDetail
    Set cross1 = crossingList(0)
    Set cross2 = crossingList(1)
    Dim pbr1 As CPBRData: Set pbr1 = ConstructPBRDataFromCrossingPairs(cross1, cross2)
    
    pbrList.Add pbr1

End Sub

Private Sub ConstructPBRInterval3(crossingList As ArrayList, pbrList As ArrayList)

    
    Dim cross1 As CTrajectoryCrossingDetail, cross2 As CTrajectoryCrossingDetail, cross3 As CTrajectoryCrossingDetail
    
    Set cross1 = crossingList(0)
    Set cross2 = crossingList(1)
    Set cross3 = crossingList(2)
    
    Dim pbr1 As CPBRData: Set pbr1 = ConstructPBRDataFromZeroToFirstCrossing(cross1)
    Dim pbr2 As CPBRData: Set pbr2 = ConstructPBRDataFromCrossingPairs(cross2, cross3)
    
    pbrList.Add pbr1
    pbrList.Add pbr2
   
End Sub

Private Sub ConstructPBRInterval4(crossingList As ArrayList, pbrList As ArrayList)

    
    Dim cross1 As CTrajectoryCrossingDetail, cross2 As CTrajectoryCrossingDetail, cross3 As CTrajectoryCrossingDetail, cross4 As CTrajectoryCrossingDetail
    
    Set cross1 = crossingList(0)
    Set cross2 = crossingList(1)
    Set cross3 = crossingList(2)
    Set cross4 = crossingList(3)
    
    Dim pbr1 As CPBRData: Set pbr1 = ConstructPBRDataFromCrossingPairs(cross1, cross2)
    Dim pbr2 As CPBRData: Set pbr2 = ConstructPBRDataFromCrossingPairs(cross3, cross4)
    
    pbrList.Add pbr1
    pbrList.Add pbr2
   
End Sub


Private Function ResolveCrossingPoint(crossPoint As CTrajectoryCrossingDetail, ByRef xRange As Double, ByRef elevationY As Double) As Boolean

    If crossPoint.ContainsTargetY(m_yMin) Then
        crossPoint.RangeAtTargetY m_yMin, xRange
        elevationY = m_yMin
        ResolveCrossingPoint = True
        Exit Function
    End If
    
    If crossPoint.ContainsTargetY(m_yMax) Then
        crossPoint.RangeAtTargetY m_yMax, xRange
        elevationY = m_yMax
        ResolveCrossingPoint = True
        Exit Function
    End If
    
    ResolveCrossingPoint = False

End Function

Private Function ResolveCrossingPointEnergy(crossPoint As CTrajectoryCrossingDetail, ByRef xRange As Double, ByRef elevationY As Double, ByRef energy As Double) As Boolean

    If crossPoint.ContainsTargetY(m_yMin) Then
        crossPoint.RangeAtTargetY m_yMin, xRange
        elevationY = m_yMin
        energy = ResolveEnergyAtRange(xRange, crossPoint)
        ResolveCrossingPointEnergy = True
        Exit Function
    End If
    
    If crossPoint.ContainsTargetY(m_yMax) Then
        crossPoint.RangeAtTargetY m_yMax, xRange
        elevationY = m_yMax
        energy = ResolveEnergyAtRange(xRange, crossPoint)
        ResolveCrossingPointEnergy = True
        Exit Function
    End If
    
    ResolveCrossingPointEnergy = False

End Function

Private Function ResolveEnergyAtRange(ByVal x As Double, crossPoint As CTrajectoryCrossingDetail) As Double

    Dim KE1 As Double: KE1 = m_trajectory(crossPoint.Index1, 3)
    Dim x1 As Double: x1 = m_trajectory(crossPoint.Index1, 1)
    Dim KE2 As Double: KE2 = m_trajectory(crossPoint.Index2, 3)
    Dim x2 As Double: x2 = m_trajectory(crossPoint.Index2, 1)
    
    Dim e1 As Double: e1 = (KE2 - KE1) / (x2 - x1)
    Dim e0 As Double: e0 = (x2 * KE1 - x1 * KE2) / (x2 - x1)
    ResolveEnergyAtRange = e1 * x + e0

End Function


Private Function ConstructPBRDataFromZeroToFirstCrossing(cross1 As CTrajectoryCrossingDetail) As CPBRData

    Dim pbr1 As New CPBRData
    
    ' min range will be at trajectory min
    Set pbr1.MinimumRange = DimensionedVariables.LengthInstance(m_trajectory(1, 1), DimensionedVariables.meters)
    Set pbr1.ElevationAtMinimumRange = DimensionedVariables.LengthInstance(m_trajectory(1, 2) / 100#, DimensionedVariables.meters)
    Set pbr1.EnergyAtMinimumRange = DimensionedVariables.EnergyInstance(m_trajectory(1, 3), DimensionedVariables.Joules)
    
    Dim maxRange As Double
    Dim elevationY As Double
    Dim energy As Double
    
    If ResolveCrossingPointEnergy(cross1, maxRange, elevationY, energy) Then
    
        Set pbr1.MaximumRange = DimensionedVariables.LengthInstance(maxRange, DimensionedVariables.meters)
        Set pbr1.ElevationAtMaximumRange = DimensionedVariables.LengthInstance(elevationY / 100#, DimensionedVariables.meters)
        Set pbr1.EnergyAtMaximumRange = DimensionedVariables.EnergyInstance(energy, DimensionedVariables.Joules)
        
    Else
    
        Set pbr1.MaximumRange = DimensionedVariables.LengthInstance(m_trajectory(m_nPoints, 1), DimensionedVariables.meters)
        Set pbr1.ElevationAtMaximumRange = DimensionedVariables.LengthInstance(m_trajectory(m_nPoints, 2) / 100#, DimensionedVariables.meters)
        Set pbr1.EnergyAtMaximumRange = DimensionedVariables.EnergyInstance(m_trajectory(m_nPoints, 3), DimensionedVariables.Joules)
        
    End If
    
    Set ConstructPBRDataFromZeroToFirstCrossing = pbr1

End Function

Private Function ConstructPBRDataFromCrossingPairs(cross1 As CTrajectoryCrossingDetail, cross2 As CTrajectoryCrossingDetail) As CPBRData

    Dim pbr1 As New CPBRData
    
    Dim minRange As Double, elevationAtMinRange As Double, energyAtMinRange As Double
    Dim maxRange As Double, elevationAtMaxRange As Double, energyAtMaxRange As Double
    
    If ResolveCrossingPointEnergy(cross1, minRange, elevationAtMinRange, energyAtMinRange) Then
        Set pbr1.MinimumRange = LengthInstance(minRange, DimensionedVariables.meters)
        Set pbr1.ElevationAtMinimumRange = LengthInstance(elevationAtMinRange / 100#, DimensionedVariables.meters)
        Set pbr1.EnergyAtMinimumRange = EnergyInstance(energyAtMinRange, DimensionedVariables.Joules)
    Else
        Set pbr1.MinimumRange = DimensionedVariables.LengthInstance(m_trajectory(1, 1), DimensionedVariables.meters)
        Set pbr1.ElevationAtMinimumRange = DimensionedVariables.LengthInstance(m_trajectory(1, 2) / 100#, DimensionedVariables.meters)
         Set pbr1.EnergyAtMinimumRange = DimensionedVariables.EnergyInstance(m_trajectory(1, 3), DimensionedVariables.Joules)
    End If
    
    If ResolveCrossingPointEnergy(cross2, maxRange, elevationAtMaxRange, energyAtMaxRange) Then
        Set pbr1.MaximumRange = LengthInstance(maxRange, DimensionedVariables.meters)
        Set pbr1.ElevationAtMaximumRange = LengthInstance(elevationAtMaxRange / 100#, DimensionedVariables.meters)
        Set pbr1.EnergyAtMaximumRange = DimensionedVariables.EnergyInstance(energyAtMaxRange, DimensionedVariables.Joules)
    Else
        Set pbr1.MaximumRange = DimensionedVariables.LengthInstance(m_trajectory(m_nPoints, 1), DimensionedVariables.meters)
        Set pbr1.ElevationAtMaximumRange = DimensionedVariables.LengthInstance(m_trajectory(m_nPoints, 2) / 100#, DimensionedVariables.meters)
        Set pbr1.EnergyAtMaximumRange = DimensionedVariables.EnergyInstance(m_trajectory(m_nPoints, 3), DimensionedVariables.Joules)
    End If
    
    Set ConstructPBRDataFromCrossingPairs = pbr1

End Function

