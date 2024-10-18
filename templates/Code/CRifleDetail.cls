VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRifleDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CoreLibrary.IGenericObject

Private m_key As String
Private m_cartridgeName As String
Private m_model As String
Private m_actionType As String
Private m_weight As Double
Private m_weightUnits As String
Private m_barrelLength As Double
Private m_overallLength As Double
Private m_lengthOfPull As Double
Private m_twistRate As Double
Private m_lengthUnits As String
Private m_sigma1Precision As Double
Private m_sigma2Precision As Double
Private m_precisionUnits As String
Private m_resolutionOuter As Double
Private m_resolutionInner As Double
Private m_resolutionUnits As String

Private m_dWeight As Weight
Private m_dBarrelLength As Length
Private m_dOverallLength As Length


Public Function Initialize() As CRifleDetail

    Set Initialize = Me

End Function


Private Property Get IGenericObject_ClassName() As String

    IGenericObject_ClassName = TypeName(Me)
    

End Property

Private Property Get IGenericObject_Key() As String

    IGenericObject_Key = m_key

End Property
