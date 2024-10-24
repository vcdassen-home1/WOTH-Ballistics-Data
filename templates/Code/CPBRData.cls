VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPBRData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description("class to encapsulate point-blank-range intervals")

Option Explicit

Private m_minRange As DimensionedVariables.Length
Private m_minRangeEnergy As DimensionedVariables.energy
Private m_maxRange As DimensionedVariables.Length
Private m_maxRangeEnergy As DimensionedVariables.energy
Private m_minElevation As DimensionedVariables.Length
Private m_maxElevation As DimensionedVariables.Length


Private Sub Class_Initialize()

    Set m_minRange = DimensionedVariables.LengthInstance(0#, DimensionedVariables.meters)
    Set m_minRangeEnergy = DimensionedVariables.EnergyInstance(0#, DimensionedVariables.Joules)
    Set m_maxRange = DimensionedVariables.LengthInstance(0#, DimensionedVariables.meters)
    Set m_maxRangeEnergy = DimensionedVariables.EnergyInstance(0#, DimensionedVariables.Joules)
    Set m_minElevation = DimensionedVariables.LengthInstance(0#, DimensionedVariables.meters)
    Set m_maxElevation = DimensionedVariables.LengthInstance(0#, DimensionedVariables.meters)
    
End Sub

Private Sub Class_Terminate()

    Set m_minRange = Nothing
    Set m_minRangeEnergy = Nothing
    Set m_maxRange = Nothing
    Set m_maxRangeEnergy = Nothing
    Set m_minElevation = Nothing
    Set m_maxElevation = Nothing

End Sub

Public Function Initialize() As CPBRData
    Set Initialize = Me
End Function

Public Property Get MinimumRange() As DimensionedVariables.Length
    Set MinimumRange = m_minRange
End Property

Public Property Set MinimumRange(arg As DimensionedVariables.Length)
    Set m_minRange = arg
End Property

Public Property Get MaximumRange() As DimensionedVariables.Length
    Set MaximumRange = m_maxRange
End Property

Public Property Set MaximumRange(arg As DimensionedVariables.Length)
    Set m_maxRange = arg
End Property

Public Property Get ElevationAtMinimumRange() As DimensionedVariables.Length
    Set ElevationAtMinimumRange = m_minElevation
End Property

Public Property Set ElevationAtMinimumRange(arg As DimensionedVariables.Length)
    Set m_minElevation = arg
End Property

Public Property Get ElevationAtMaximumRange() As DimensionedVariables.Length
    Set ElevationAtMaximumRange = m_maxElevation
End Property

Public Property Set ElevationAtMaximumRange(arg As DimensionedVariables.Length)
    Set m_maxElevation = arg
End Property

Public Property Get EnergyAtMinimumRange() As DimensionedVariables.energy
    Set EnergyAtMinimumRange = m_minRangeEnergy
End Property

Public Property Set EnergyAtMinimumRange(arg As DimensionedVariables.energy)
    Set m_minRangeEnergy = arg
End Property

Public Property Get EnergyAtMaximumRange() As DimensionedVariables.energy
    Set EnergyAtMaximumRange = m_maxRangeEnergy
End Property

Public Property Set EnergyAtMaximumRange(arg As DimensionedVariables.energy)
    Set m_maxRangeEnergy = arg
End Property

