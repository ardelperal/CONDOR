VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Interfaz: ISolicitud
' Descripción: Interfaz base para todas las clases de solicitudes en CONDOR
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Propiedades comunes a todas las solicitudes
' Basadas en la especificación funcional del sistema CONDOR

' Identificador único de la solicitud
Public Property Get ID_Solicitud() As Long
End Property

' Número del expediente asociado
Public Property Get ID_Expediente() As String
End Property

' Tipo de solicitud (PC, CD_CA, CD_CA_SUB)
Public Property Get TipoSolicitud() As String
End Property

' Código único de la solicitud
Public Property Get CodigoSolicitud() As String
End Property

' Estado interno de la solicitud en el workflow
Public Property Get EstadoInterno() As String
End Property

' Métodos comunes a todas las solicitudes

' Carga los datos de una solicitud por su ID
' Parámetros:
'   id: Identificador único de la solicitud
' Retorna: True si la carga fue exitosa, False en caso contrario
Public Function Load(ByVal id As Long) As Boolean
End Function

' Guarda los datos de la solicitud en la base de datos
' Retorna: True si el guardado fue exitoso, False en caso contrario
Public Function Save() As Boolean
End Function

' Cambia el estado de la solicitud siguiendo las reglas del workflow
' Parámetros:
'   newState: Nuevo estado al que se desea cambiar
' Retorna: True si el cambio de estado fue exitoso, False en caso contrario
Public Function ChangeState(ByVal newState As String) As Boolean
End Function