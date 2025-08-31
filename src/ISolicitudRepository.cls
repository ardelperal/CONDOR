VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =====================================================
' INTERFAZ: ISolicitudRepository
' DESCRIPCIÓN: Define el contrato para el repositorio de solicitudes.
'              Proporciona métodos para obtener y guardar solicitudes.
' AUTOR: CONDOR-Architect
' FECHA: 2025-08-28
' =====================================================

Public Function GetSolicitudById(ByVal idSolicitud As Long) As ESolicitud
End Function

Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
End Function