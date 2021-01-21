VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_enum As IEnumVARIANT

Public Property Set EnumVariant(newEnum_ As IEnumVARIANT)
    Set m_enum = newEnum_
End Property
Public Property Get EnumVariant() As IEnumVARIANT
Attribute EnumVariant.VB_UserMemId = -4
    Set EnumVariant = m_enum
End Property

Public Property Get Self() As EnumHelper
    Set Self = Me
End Property
