Attribute VB_Name = "UnitsConverters"
Option Explicit


Public Const lengthUnitsIn As String = "in"
Public Const lengthUnitFt As String = "ft"
Public Const lengthUnitYard As String = "yd"
Public Const lengthUnitMile As String = "mi"

Public Const lengthUnitMM As String = "mm"
Public Const lengthUnitCM As String = "cm"
Public Const lengthUnitMeter As String = "m"
Public Const lengthUnitKm As String = "km"

Public Const weightUnitGrain As String = "grain"
Public Const weightUnitLb As String = "lbm"
Public Const massUnitGrams As String = "g"
Public Const massUnitKilos As String = "kg"

Public Const speedUnitFPS As String = "ft/sec"
Public Const speedUnitMPH As String = "mi/hr"
Public Const speedUnitMPS As String = "m/sec"
Public Const speedUnitKPH As String = "km/hr"

Public Const energyUnitFLB As String = "flb"
Public Const energyUnitJoules As String = "J"

Public Function ConvertLength(ByVal value As Double, ByVal fromUnits As String, ByVal toUnits As String) As Double

    ConvertLength = Application.WorksheetFunction.Convert(value, fromUnits, toUnits)

End Function

Public Function ConvertSpeed(ByVal value As Double, ByVal fromUnits As String, ByVal toUnits As String) As Double

    Dim convertedValue As Double
    
    If fromUnits = speedUnitFPS And toUnits = speedUnitMPS Then
        convertedValue = Application.WorksheetFunction.Convert(value, lengthUnitFt, lengthUnitMeter)
    ElseIf fromUnits = speedUnitMPS And toUnits = speedUnitFPS Then
        convertedValue = Application.WorksheetFunction.Convert(value, lengthUnitMeter, lengthUnitFt)
    ElseIf fromUnits = toUnits Then
        convertedValue = value
    Else
        convertedValue = Application.WorksheetFunction.Convert(value, fromUnits, toUnits)
    End If
    
    ConvertSpeed = convertedValue
        
End Function

Public Function ConvertMassWeight(ByVal value As Double, ByVal fromUnits As String, ByVal toUnits As String) As Double
    ConvertMassWeight = Application.WorksheetFunction.Convert(value, fromUnits, toUnits)
End Function



