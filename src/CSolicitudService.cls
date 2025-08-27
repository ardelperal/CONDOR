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
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACIÃ“N Y CONFIGURACIÃ“N
' ============================================================================

' Inicializa el servicio con sus dependencias
' @param solicitudRepository: Repositorio para operaciones de persistencia
' @param operationLogger: Logger para registrar operaciones
Public Sub Initialize(ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger)
    If solicitudRepository Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "solicitudRepository no puede ser Nothing"
    End If
    
    If operationLogger Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "operationLogger no puede ser Nothing"
    End If
    
    Set m_solicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE ISolicitudService
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
Private Function ISolicitudService_CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As T_Solicitud
    Dim ErrorHandler As IErrorHandlerService
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If Len(Trim(idExpediente)) = 0 Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "idExpediente no puede estar vacío"
    End If
    
    If Len(Trim(tipo)) = 0 Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "tipo no puede estar vacío"
    End If
    
    On Error GoTo ErrorHandler
    
    ' Crear nueva solicitud con estado inicial "Borrador"
    Dim nuevaSolicitud As T_Solicitud
    Set nuevaSolicitud = New T_Solicitud
    
    With nuevaSolicitud
        .idSolicitud = 0 ' Será asignado por el repositorio
        .idExpediente = Trim(idExpediente)
        .tipoSolicitud = Trim(tipo)
        .subTipoSolicitud = ""
        .codigoSolicitud = GenerateCodigoSolicitud(idExpediente, tipo)
        .idEstadoInterno = 1 ' ID del estado Borrador en tbEstados
        .fechaCreacion = Now()
        .usuarioCreacion = Environ("USERNAME") ' Usuario actual del sistema
        .fechaModificacion = Null
        .usuarioModificacion = ""
    End With
    
    ' Delegar al repositorio para guardar
    Dim nuevoId As Long
    nuevoId = m_solicitudRepository.SaveSolicitud(nuevaSolicitud)
    
    If nuevoId > 0 Then
        nuevaSolicitud.idSolicitud = nuevoId
        m_OperationLogger.LogOperation "CREATE", CStr(nuevoId), _
            "Nueva solicitud creada: ID=" & nuevoId & ", Expediente=" & idExpediente & ", Tipo=" & tipo
        
        Set ISolicitudService_CreateSolicitud = nuevaSolicitud
    Else
        m_OperationLogger.LogOperation "CREATE_ERROR", CStr(idExpediente), _
            "Error: No se pudo crear la solicitud para expediente " & idExpediente
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "No se pudo crear la solicitud"
    End If
    
    Exit Function
    
ErrorHandler:
    m_OperationLogger.LogOperation "CREATE_ERROR", "", "Error: " & Err.Description
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.CreateSolicitud"
    Err.Raise Err.Number, "CSolicitudService.CreateSolicitud", Err.Description
End Function

' Guarda una solicitud existente
Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As T_Solicitud) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If solicitud Is Nothing Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "solicitud no puede ser Nothing"
    End If
    
    On Error GoTo ErrorHandler
    
    ' Actualizar campos de modificación
    solicitud.fechaModificacion = Now()
    solicitud.usuarioModificacion = Environ("USERNAME")
    
    ' Delegar al repositorio para guardar
    Dim resultId As Long
    resultId = m_solicitudRepository.SaveSolicitud(solicitud)
    
    If resultId > 0 Then
        m_OperationLogger.LogOperation "SAVE", CStr(solicitud.idSolicitud), _
            "Solicitud guardada correctamente: ID=" & solicitud.idSolicitud
        ISolicitudService_SaveSolicitud = True
    Else
        m_OperationLogger.LogOperation "SAVE_ERROR", CStr(solicitud.idSolicitud), _
            "Error: No se pudo guardar la solicitud ID=" & solicitud.idSolicitud
        ISolicitudService_SaveSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    m_OperationLogger.LogOperation "SAVE_ERROR", "", "Error: " & Err.Description
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.SaveSolicitud"
    ISolicitudService_SaveSolicitud = False
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
Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String) As T_Solicitud
    Set CreateSolicitud = ISolicitudService_CreateSolicitud(idExpediente, tipo)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Boolean
    SaveSolicitud = ISolicitudService_SaveSolicitud(solicitud)
End Function














