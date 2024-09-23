VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCartridgeLoadoutDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_ammunitionTableName As String
Private m_ammunitionData As IVariantArray2D
Private m_rifleTableName As String
Private m_rifleData As IVariantArray2D

Public Function Initialize( _
    ByVal ammoTblName As String, _
    ammoData2D As IVariantArray2D, _
    ByVal rifleTblName As String, _
    rifleData2D As IVariantArray2D) As CCartridgeLoadoutDetail
    
    m_ammunitionTableName = ammoTblName
    Set m_ammunitionData = ammoData2D
    
    m_rifleTableName = rifleTblName
    Set m_rifleData = rifleData2D
    
    Set Initialize = Me
    
End Function



Private Sub Class_Initialize()

    Set m_ammunitionData = Nothing
    Set m_rifleData = Nothing

End Sub

Private Sub Class_Terminate()

    Set m_ammunitionData = Nothing
    Set m_rifleData = Nothing

End Sub

Public Property Get AmmunitionTableName() As String
    AmmunitionTableName = m_ammunitionTableName
End Property

Public Property Get ammunitionData() As IVariantArray2D
    Set ammunitionData = m_ammunitionData

End Property

Public Property Get RifleTableName() As String
    RifleTableName = m_rifleTableName
End Property

Public Property Get rifleData() As IVariantArray2D
    Set rifleData = m_rifleData

End Property
