VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' INTERFAZ: ISolicitudRepository
' DESCRIPCIÓN: Define el contrato para la persistencia de solicitudes PC
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

Option Explicit

'******************************************************************************
' MÉTODOS DE PERSISTENCIA
'******************************************************************************

' Guarda una solicitud PC y retorna el ID asignado
Public Function SavePC(solicitud As CSolicitudPC) As Long
End Function

' Carga una solicitud PC por su ID
Public Function LoadPC(ByVal ID As Long) As CSolicitudPC
End Function