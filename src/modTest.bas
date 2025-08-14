Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit


' Módulo de prueba para verificar la implementación de la interfaz
Public Sub TestInterface()
    On Error GoTo ErrorHandler
    
    Dim solicitud As ISolicitud
    Dim solicitudPC As CSolicitudPC
    
    ' Crear instancia de CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Asignar a la interfaz
    Set solicitud = solicitudPC
    
    ' Probar propiedades
    solicitud.ID_Solicitud = 1
    solicitud.ID_Expediente = "EXP-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.EstadoInterno = "Borrador"
    
    Debug.Print "Test exitoso: ID=" & solicitud.ID_Solicitud
    Debug.Print "Test exitoso: Expediente=" & solicitud.ID_Expediente
    Debug.Print "Test exitoso: Tipo=" & solicitud.TipoSolicitud
    Debug.Print "Test exitoso: Estado=" & solicitud.EstadoInterno
    
    ' Limpiar objetos
    Set solicitud = Nothing
    Set solicitudPC = Nothing
    
    Debug.Print "Prueba de interfaz completada exitosamente"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en prueba de interfaz: " & Err.Description
    Set solicitud = Nothing
    Set solicitudPC = Nothing
End Sub

