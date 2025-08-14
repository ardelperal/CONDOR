Attribute VB_Name = "modSolicitudFactory"
' ============================================================================
' Módulo: modSolicitudFactory
' Descripción: Factory Pattern para crear instancias de solicitudes
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DEL FACTORY
' ============================================================================

' Función que crea una instancia de solicitud basada en el ID
' Por ahora retorna siempre CSolicitudPC, la lógica completa se implementará después
Public Function CreateSolicitud(ByVal idSolicitud As Long) As ISolicitud
    On Error GoTo ErrorHandler
    
    ' TODO: Implementar lógica para determinar el tipo de solicitud
    ' basándose en la consulta a Tb_Solicitudes
    ' Por ahora, siempre crea una CSolicitudPC para que compile
    
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Cargar los datos de la solicitud
    If solicitud.Load(idSolicitud) Then
        Set CreateSolicitud = solicitud
    Else
        Set CreateSolicitud = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    Set CreateSolicitud = Nothing
End Function

' ============================================================================
' FUNCIONES AUXILIARES (PARA IMPLEMENTACIÓN FUTURA)
' ============================================================================

' Función auxiliar para determinar el tipo de solicitud
' TODO: Implementar consulta a Tb_Solicitudes para obtener TipoSolicitud
Private Function GetTipoSolicitud(ByVal idSolicitud As Long) As String
    ' Por ahora retorna "PC" por defecto
    GetTipoSolicitud = "PC"
End Function

' Función para crear solicitud de tipo PC (Propuesta de Cambio)
Private Function CreateSolicitudPC(ByVal idSolicitud As Long) As ISolicitud
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    If solicitud.Load(idSolicitud) Then
        Set CreateSolicitudPC = solicitud
    Else
        Set CreateSolicitudPC = Nothing
    End If
End Function

' TODO: Agregar funciones para otros tipos de solicitud cuando se implementen
' Private Function CreateSolicitudCD_CA(ByVal idSolicitud As Long) As ISolicitud
' Private Function CreateSolicitudCD_CA_SUB(ByVal idSolicitud As Long) As ISolicitud
