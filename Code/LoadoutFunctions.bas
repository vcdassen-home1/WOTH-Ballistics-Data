Attribute VB_Name = "LoadoutFunctions"
Option Explicit

Public Enum ERifleTableColumns

    eKey = 1
    eChambering = 2
    eModel = 3
    eAction = 4
    eWeight = 5
    eWeightUnits = 6
    eBarrelLength = 7
    eOverallLength = 8
    eTwistRate = 9
    eLengthUnits = 10
    ePrecision1Sigma = 11
    ePrecision2Sigma = 12
    ePrecisionUnits = 13
    eResolutionOuter = 14
    eResolutionInner = 15
    eResolutionUnits = 16
    
End Enum

Public Enum EAmmunitionTableColumns

    eKey = 1
    eSpecificationType = 2
    eCartridge = 3
    eProjectile = 4
    eType = 5
    eCaliber = 6
    eCaliberUnits = 7
    eWeight = 8
    eWeightUnits = 9
    eMuzzleVelocity = 10
    eMuzzleVelocityUnits = 11
    eMinimumTerminalVelocity = 12
    eMinimumTerminalVelocityUnits = 13
    eBCG1 = 14
    eBCG7 = 15
    eCaliberInches = 16
    eCaliberMM = 17
    eWeightGrains = 18
    eWeightGrams = 19
    eMuzzleVelocityFeetPerSecond = 20
    eMuzzleVelocityMetersPerSecond = 21
    eMuzzleEnergyFtLbs = 22
    eMuzzleEnergyJoules = 23
    eMinimumTerminalVelocityFeetPerSecond = 24
    eMinimumTerminalVelocityMetersPerSecond = 25
    
End Enum

Public Const AmmunitionMfr As String = "Manufacturer"
Public Const AmmunitionGame As String = "Game"


Public Function CartridgeLoadoutDetail(ByVal cartridgeName As String) As CCartridgeLoadoutDetail

    If CoreLibrary.IsNotEmpty(cartridgeName) = False Then
        Set CartridgeLoadoutDetail = Nothing
        Exit Function
    End If

    Dim cartridgeData As IVariantArray2D: Set cartridgeData = LoadCartridgeData()
    
    If cartridgeData Is Nothing Then
        Set CartridgeLoadoutDetail = Nothing
        Exit Function
    End If
    
    Dim ammoTableName As String: ammoTableName = AmmunitionTableName(cartridgeData, cartridgeName)
    Dim rifleTblName As String: rifleTblName = RifleTableName(cartridgeData, cartridgeName)
    If CoreLibrary.IsNotEmpty(ammoTableName) = False Or CoreLibrary.IsNotEmpty(rifleTblName) = False Then
        Set CartridgeLoadoutDetail = Nothing
        Exit Function
    End If
    
    Dim ammunitionData As IVariantArray2D: Set ammunitionData = LoadTableData(ammoTableName, shLoadouts)
    If ammunitionData Is Nothing Then
        Set CartridgeLoadoutDetail = Nothing
        Exit Function
    End If
        
    
    Dim rifleData As IVariantArray2D: Set rifleData = LoadTableData(rifleTblName, shLoadouts)
    If rifleData Is Nothing Then
        Set CartridgeLoadoutDetail = Nothing
        Exit Function
    End If
    
    Dim loadoutData As New CCartridgeLoadoutDetail
    
    Set CartridgeLoadoutDetail = loadoutData.Initialize(ammoTableName, ammunitionData, rifleTblName, rifleData)
    
End Function

Public Function RifleNameList(rifleData2D As IVariantArray2D) As ArrayList
    Set RifleNameList = ExtractColumn(rifleData2D, ERifleTableColumns.eModel)
End Function

Public Function AmmunitionNameList(ammoData2D As IVariantArray2D, ByVal spec As String) As ArrayList

    Dim matchingRows As ArrayList: Set matchingRows = FindMatchingRows(ammoData2D, spec, EAmmunitionTableColumns.eSpecificationType)
    Set AmmunitionNameList = ExtractRowColumns(ammoData2D, matchingRows, EAmmunitionTableColumns.eProjectile)

End Function







