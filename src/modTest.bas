Option Compare Database
Option Explicit
' M?dulo de prueba para verificar la implementaci?n de la interfaz
Public Sub TestInterface()
    On Error GoTo ErrorHandler
    
    Dim solicitud As ISolicitud
    Dim solicitudPC As CSolicitudPC
    
    ' Crear instancia de CSolicitudPC
    Set solicitudPC = New CSolicitudPC
    
    ' Asignar a la interfaz
    Set solicitud = solicitudPC
    
    ' Probar propiedades
    solicitud.idSolicitud = 1
    solicitud.idExpediente = "EXP-001"
    solicitud.tipoSolicitud = "PC"
    solicitud.estadoInterno = "Borrador"
    
    Debug.Print "Test exitoso: ID=" & solicitud.idSolicitud
    Debug.Print "Test exitoso: Expediente=" & solicitud.idExpediente
    Debug.Print "Test exitoso: Tipo=" & solicitud.tipoSolicitud
    Debug.Print "Test exitoso: Estado=" & solicitud.estadoInterno
    
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











