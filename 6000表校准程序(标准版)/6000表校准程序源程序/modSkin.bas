Attribute VB_Name = "modSkin"
Public Declare Function SkinH_Attach Lib "SkinH_VB6.dll" () As Long

Public Declare Function SkinH_AttachEx Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
