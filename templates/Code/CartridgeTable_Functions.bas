Attribute VB_Name = "CartridgeTable_Functions"
Option Explicit

Public Enum ECartridgeTableColumns

    eChamberings = 1
    eBulletCaliber = 2
    eCaliberUnits = 3
    eAmmunitionTable = 4
    eRifleTable = 5
    
End Enum


Public Function LoadCartridgeData() As IVariantArray2D

    Set LoadCartridgeData = LoadTableData("CartridgeNames", shConstants)
   
End Function

Public Function AmmunitionTableName(cartridgeData As IVariantArray2D, ByVal cartridgeName As String) As String

    Dim row As Long: row = FindMatchingRow(cartridgeData, cartridgeName, ECartridgeTableColumns.eChamberings)
    If cartridgeData.rowExtent.IsIndexInRange(row) Then
        AmmunitionTableName = cartridgeData.Element(CoreLibrary.Array2DIndex(row, ECartridgeTableColumns.eAmmunitionTable))
        Exit Function
    End If
    
    AmmunitionTableName = ""
    

End Function

Public Function RifleTableName(cartridgeData As IVariantArray2D, ByVal cartridgeName As String) As String

    Dim row As Long: row = FindMatchingRow(cartridgeData, cartridgeName, ECartridgeTableColumns.eChamberings)
    If cartridgeData.rowExtent.IsIndexInRange(row) Then
        RifleTableName = cartridgeData.Element(CoreLibrary.Array2DIndex(row, ECartridgeTableColumns.eRifleTable))
        Exit Function
    End If
    
    RifleTableName = ""
    
End Function

