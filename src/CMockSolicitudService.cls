VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockSolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Mock del Servicio de Solicitudes para Pruebas
' Implementa ISolicitudService con datos hardcodeados y predecibles
' Version: 1.0
' Fecha: 2025-01-14


Implements ISolicitudService

' Mock: Crear una nueva solicitud
Private Function ISolicitudService_CreateNuevaSolicitud(ByVal tipo As String) As ISolicitud
    ' Retorna un mock de solicitud con datos predecibles
    Dim mockSolicitud As New CSolicitudPC
    
    ' Configurar datos mock seg?n el tipo
    Select Case UCase(tipo)
        Case "PC"
            ' Configurar solicitud PC mock
            ' TODO: Implementar configuraci?n espec?fica cuando CSolicitudPC est? completa
        Case Else
            ' Tipo por defecto
    End Select
    
    Set ISolicitudService_CreateNuevaSolicitud = mockSolicitud
End Function

' Mock: Obtener una solicitud por su ID
Private Function ISolicitudService_GetSolicitudPorID(ByVal ID As Long) As ISolicitud
    ' Retorna un mock de solicitud con ID espec?fico
    Dim mockSolicitud As New CSolicitudPC
    
    ' Configurar datos mock seg?n el ID
    Select Case ID
        Case 1
            ' Solicitud de prueba #1
        Case 2
            ' Solicitud de prueba #2
        Case Else
            ' ID no encontrado - retorna Nothing
            Set ISolicitudService_GetSolicitudPorID = Nothing
            Exit Function
    End Select
    
    Set ISolicitudService_GetSolicitudPorID = mockSolicitud
End Function

' Mock: Guardar una solicitud
Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As ISolicitud) As Boolean
    ' Simula guardado exitoso si la solicitud no es Nothing
    If Not solicitud Is Nothing Then
        ISolicitudService_SaveSolicitud = True
    Else
        ISolicitudService_SaveSolicitud = False
    End If
End Function

' Mock: Obtener todas las solicitudes
Private Function ISolicitudService_GetAllSolicitudes() As Collection
    ' Retorna una colecci?n con solicitudes mock
    Dim coleccion As New Collection
    Dim solicitud1 As New CSolicitudPC
    Dim solicitud2 As New CSolicitudPC
    
    ' Agregar solicitudes mock a la colecci?n
    coleccion.Add solicitud1
    coleccion.Add solicitud2
    
    Set ISolicitudService_GetAllSolicitudes = coleccion
End Function

' Mock: Eliminar una solicitud
Private Function ISolicitudService_DeleteSolicitud(ByVal ID As Long) As Boolean
    ' Simula eliminaci?n exitosa para IDs v?lidos
    If ID > 0 And ID <= 100 Then
        ISolicitudService_DeleteSolicitud = True
    Else
        ISolicitudService_DeleteSolicitud = False
    End If
End Function

' Mock: Actualizar el estado de una solicitud
Private Function ISolicitudService_UpdateEstadoSolicitud(ByVal ID As Long, ByVal nuevoEstado As String) As Boolean
    ' Simula actualizaci?n exitosa para IDs y estados v?lidos
    If ID > 0 And ID <= 100 And Len(nuevoEstado) > 0 Then
        ISolicitudService_UpdateEstadoSolicitud = True
    Else
        ISolicitudService_UpdateEstadoSolicitud = False
    End If
End Function


