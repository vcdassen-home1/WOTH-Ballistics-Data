Attribute VB_Name = "Ballistics_CodeBehind"
Option Explicit

Private Const rifleSelectorCellAddress As String = "$B$5"
Private Const manufacturerAmmoSelectorCellAddress As String = "$B$15"
Private Const gameAmmoSelectorCellAddress As String = "$B$23"
Private Const actualAmmoSelectorCellAddress As String = "$B$31"

Private Const rifleTableRegionAddress As String = "$A$4:$F$11"
Private Const manufacturerAmmoTableRegionAddress As String = "$A$14:$F$20"
Private Const manufacturerAmmoTableBarrelCorrectionsRegionAddress = "$G$14:$J$20"
Private Const gameAmmoTableRegionAddress As String = "$A$22:$F$28"
Private Const actualAmmoTableRegionAddress As String = "$A$30:$F$36"


Public Sub PopulateCartridgeLoadouts(ws As Worksheet, cartridgeDetail As CCartridgeLoadoutDetail)

    Dim cellProperties As CoreLibrary.CCellProperties: Set cellProperties = CoreLibrary.CellPropertiesClass()
    SetRifleLoadoutSelector ws, cellProperties, cartridgeDetail
    SetAmmunitionLoadoutSelector ws, cellProperties, cartridgeDetail
    
End Sub

Public Sub ClearCartridgeLoadouts(ws As Worksheet)

   ClearRifleLoadout ws
   ClearAmmunitionLoadouts ws, True

End Sub

Private Sub SetRifleLoadoutSelector(ws As Worksheet, cellProperties As CoreLibrary.CCellProperties, cartridgeDetail As CCartridgeLoadoutDetail)

    Dim rifleSelectorCell As Range: Set rifleSelectorCell = ws.Range(rifleSelectorCellAddress)
    Dim rifleNames As ArrayList: Set rifleNames = RifleNameList(cartridgeDetail.rifleData)
    cellProperties.SetListValidationPropertiesForCell rifleSelectorCell, rifleNames, "Available Rifles", bSetValue:=True
    rifleSelectorCell.value = rifleNames(0)

End Sub

Private Sub ClearRifleLoadout(ws As Worksheet)

    Dim rifleSelectorCell As Range: Set rifleSelectorCell = ws.Range(rifleSelectorCellAddress)
    rifleSelectorCell.Validation.Delete
    rifleSelectorCell.value = ""


End Sub


Private Sub SetAmmunitionLoadoutSelector(ws As Worksheet, cellProperties As CoreLibrary.CCellProperties, cartridgeDetail As CCartridgeLoadoutDetail)

    Dim ammoSelectorCell As Range: Set ammoSelectorCell = ws.Range(manufacturerAmmoSelectorCellAddress)
    Dim ammoSelectorCell2 As Range: Set ammoSelectorCell2 = ws.Range(gameAmmoSelectorCellAddress)
    Dim ammoSelectorCell3 As Range: Set ammoSelectorCell3 = ws.Range(actualAmmoSelectorCellAddress)
    
    Dim ammunitionNames As ArrayList: Set ammunitionNames = AmmunitionNameList(cartridgeDetail.ammunitionData, AmmunitionMfr)
    cellProperties.SetListValidationPropertiesForCell ammoSelectorCell, ammunitionNames, "Available Ammunition", bSetValue:=True
    ammoSelectorCell.value = ammunitionNames(0)
    ammoSelectorCell2.value = ammunitionNames(0)
    ammoSelectorCell3.value = ammunitionNames(0)

End Sub

Public Sub PopulateAmmunitionLoadouts(ws As Worksheet, ByVal projectileName As String, cartridgeDetail As CCartridgeLoadoutDetail)

    Dim gameAmmoProjectileCell As Range: Set gameAmmoProjectileCell = ws.Range(gameAmmoSelectorCellAddress)
    Dim actualAmmoProjectileCell As Range: Set actualAmmoProjectileCell = ws.Range(actualAmmoSelectorCellAddress)
    gameAmmoProjectileCell.value = projectileName
    actualAmmoProjectileCell.value = projectileName

End Sub

Public Sub ClearAmmunitionLoadouts(ws As Worksheet, Optional ByVal removeValidation As Boolean = False)

    Dim ammoSelectorCell As Range: Set ammoSelectorCell = ws.Range(manufacturerAmmoSelectorCellAddress)
    Dim ammoSelectorCell2 As Range: Set ammoSelectorCell2 = ws.Range(gameAmmoSelectorCellAddress)
    Dim ammoSelectorCell3 As Range: Set ammoSelectorCell3 = ws.Range(actualAmmoSelectorCellAddress)
    
    If removeValidation Then
        ammoSelectorCell.Validation.Delete
    End If
    ammoSelectorCell.value = ""
    ammoSelectorCell2.value = ""
    ammoSelectorCell3.value = ""

End Sub

Private Sub PopulateRifleDataTable(ws As Worksheet, ByVal rifleName As String, rifleData2D As IVariantArray2D)

End Sub

Private Sub PopulateManufacturerAmmoData(ws As Worksheet, ByVal projectileName As String, ammoData2D As IVariantArray2D)

End Sub

Private Sub PopulateManufacturerAmmoBarrelLengthCorrections(ws As Worksheet, ByVal projectileName As String, ByVal rifleName As String, cartridgeDetail As CCartridgeLoadoutDetail)

End Sub

Private Sub PopulateGameAmmoData(ws As Worksheet, ByVal projectileName As String, ammoData2D As IVariantArray2D)

End Sub


