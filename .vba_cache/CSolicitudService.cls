VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' ============================================================================
' CLASE: CSolicitudService
' Arquitectura: Implementación del Servicio de Solicitudes
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Implementa ISolicitudService con lógica de negocio y delegación al repositorio.
' Aplica principios de inyección de dependencias y separación de responsabilidades.
' ============================================================================

Implements ISolicitudService

' Dependencias inyectadas
Private m_solicitudRepository As ISolicitudRepository
Private m_OperationLogger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACIÃ“N Y CONFIGURACIÃ“N
' ============================================================================

' Inicializa el servicio con sus dependencias
' @param solicitudRepository: Repositorio para operaciones de persistencia
' @param operationLogger: Logger para registrar operaciones
' @param errorHandler: Manejador de errores para logging
Public Sub Initialize(ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger, ByVal errorHandler As IErrorHandlerService)
    If solicitudRepository Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "solicitudRepository no puede ser Nothing"
    End If
    
    If operationLogger Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "operationLogger no puede ser Nothing"
    End If
    
    If errorHandler Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "errorHandler no puede ser Nothing"
    End If
    
    Set m_solicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    Set m_ErrorHandler = errorHandler
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE ISolicitudService
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
Private Function ISolicitudService_CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If paraExpediente Is Nothing Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "paraExpediente no puede ser Nothing"
    End If
    

    
    On Error GoTo errorHandler
    
    ' Crear nueva solicitud con estado inicial "Borrador"
    Dim nuevaSolicitud As ESolicitud
    Set nuevaSolicitud = New ESolicitud
    
    With nuevaSolicitud
        .idSolicitud = 0 ' Será asignado por el repositorio
        .idExpediente = paraExpediente.idExpediente
        .tipoSolicitud = "PC" ' Tipo por defecto
        .subTipoSolicitud = ""
        .codigoSolicitud = GenerateCodigoSolicitud(CStr(paraExpediente.idExpediente), "PC")
        .idEstadoInterno = 1 ' ID del estado Borrador en tbEstados
        .fechaCreacion = Now()
        .usuarioCreacion = Environ("USERNAME") ' Usuario actual del sistema
        .fechaModificacion = Null
        .usuarioModificacion = ""
    End With
    
    ' Delegar al repositorio para guardar
    If m_solicitudRepository.SaveSolicitud(nuevaSolicitud) Then
        m_OperationLogger.LogOperation "CREATE", CStr(nuevaSolicitud.idSolicitud), _
            "Nueva solicitud creada: ID=" & nuevaSolicitud.idSolicitud & ", Expediente=" & paraExpediente.idExpediente & ", Tipo=PC"
        
        Set ISolicitudService_CreateSolicitud = nuevaSolicitud
    Else
        m_OperationLogger.LogOperation "CREATE_ERROR", CStr(paraExpediente.idExpediente), _
            "Error: No se pudo crear la solicitud para expediente " & paraExpediente.idExpediente
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "No se pudo crear la solicitud"
    End If
    
    Exit Function
    
errorHandler:
    m_OperationLogger.LogOperation "CREATE_ERROR", "", "Error: " & Err.Description
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.CreateSolicitud"
    Err.Raise Err.Number, "CSolicitudService.CreateSolicitud", Err.Description
End Function

' Guarda una solicitud existente
Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If solicitud Is Nothing Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "solicitud no puede ser Nothing"
    End If
    
    On Error GoTo errorHandler
    
    ' Actualizar campos de modificación
    solicitud.fechaModificacion = Now()
    solicitud.usuarioModificacion = Environ("USERNAME")
    
    ' Delegar al repositorio para guardar
    Dim Resultado As Long
    Resultado = m_solicitudRepository.SaveSolicitud(solicitud)
    If Resultado > 0 Then
        m_OperationLogger.LogOperation "SAVE", CStr(solicitud.idSolicitud), _
            "Solicitud guardada correctamente: ID=" & solicitud.idSolicitud
        ISolicitudService_SaveSolicitud = Resultado
    Else
        m_OperationLogger.LogOperation "SAVE_ERROR", CStr(solicitud.idSolicitud), _
            "Error: No se pudo guardar la solicitud ID=" & solicitud.idSolicitud
        ISolicitudService_SaveSolicitud = 0
    End If
    
    Exit Function
    
errorHandler:
    m_OperationLogger.LogOperation "SAVE_ERROR", "", "Error: " & Err.Description
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.SaveSolicitud"
    ISolicitudService_SaveSolicitud = 0
End Function

' ============================================================================
' MÃ‰TODOS PRIVADOS DE APOYO
' ============================================================================

' Genera un código único para la solicitud
' @param idExpediente: ID del expediente
' @param tipo: Tipo de solicitud
' @return: Código generado
Private Function GenerateCodigoSolicitud(ByVal idExpediente As String, ByVal tipo As String) As String
    Dim timestamp As String
    timestamp = Format(Now(), "yyyymmddhhnnss")
    
    GenerateCodigoSolicitud = tipo & "-" & idExpediente & "-" & timestamp
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'==============================================================================

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    Set CreateSolicitud = ISolicitudService_CreateSolicitud(paraExpediente)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudService_SaveSolicitud(solicitud)
End Function

' Obtiene una solicitud por su ID
Private Function ISolicitudService_ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    ' Asumimos que el repositorio tiene un método similar
    Set ISolicitudService_ObtenerSolicitudPorId = m_solicitudRepository.ObtenerSolicitudPorId(SolicitudID)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Set ObtenerSolicitudPorId = ISolicitudService_ObtenerSolicitudPorId(SolicitudID)
End Function

















