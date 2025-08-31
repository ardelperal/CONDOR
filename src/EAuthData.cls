VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EAuthData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Proyecto: CONDOR
' Clase: EAuthData
' Descripción: Objeto de transferencia de datos (DTO) para información de autenticación
' Autor: Sistema CONDOR
' Fecha: 2024
' =============================================================================

Option Explicit

' Propiedades públicas para datos de autenticación
Public UserExists As Boolean
Public IsGlobalAdmin As Boolean
Public IsAppAdmin As Boolean
Public IsCalidad As Boolean
Public IsTecnico As Boolean

' Constructor por defecto
Private Sub Class_Initialize()
    UserExists = False
    IsGlobalAdmin = False
    IsAppAdmin = False
    IsCalidad = False
    IsTecnico = False
End Sub