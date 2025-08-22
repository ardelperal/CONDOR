Option Compare Database
Option Explicit
' Implementación del Servicio de Solicitudes
' Centraliza toda la lógica de negocio para operaciones con solicitudes
' Arquitectura: Capa de Lógica de Negocio con Inversión de Dependencias
' Version: 2.0 - Refactorización Arquitectónica
' Fecha: 2025-01-14

Implements ISolicitudService

' Variables privadas para dependencias inyectadas
Private m_SolicitudRepository As ISolicitudRepository
Private m_OperationLogger As IOperationLogger

' Método de inicialización para inyectar dependencias
Public Sub Initialize(solicitudRepository As ISolicitudRepository, operationLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    
    Set m_SolicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.Initialize"
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE LÓGICA DE NEGOCIO
' ============================================================================

' Crea una nueva solicitud con lógica de negocio completa
Public Function CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String, ByVal subTipo As String) As T_Solicitud
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría: Inicio de creación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "CREATE_SOLICITUD_START", "Creando solicitud - Expediente: " & idExpediente & ", Tipo: " & tipo, "CSolicitudService.CreateSolicitud"
    End If
    
    ' Crear nueva instancia de T_Solicitud
    Dim nuevaSolicitud As T_Solicitud
    Set nuevaSolicitud = New T_Solicitud
    
    ' Asignar parámetros recibidos
    nuevaSolicitud.idExpediente = idExpediente
    nuevaSolicitud.tipoSolicitud = tipo
    nuevaSolicitud.subTipoSolicitud = subTipo
    
    ' Generar código único de solicitud (simulado por ahora)
    nuevaSolicitud.codigoSolicitud = "SOL-" & Format(Now(), "yyyymmddhhnnss")
    
    ' Establecer estado inicial
    nuevaSolicitud.estadoInterno = "Borrador"
    
    ' Rellenar campos de auditoría
    nuevaSolicitud.fechaCreacion = Now()
    nuevaSolicitud.usuarioCreacion = Environ("USERNAME") ' Usuario actual del sistema
    
    ' Persistir usando el repositorio
    Dim idGenerado As Long
    idGenerado = m_SolicitudRepository.GuardarSolicitud(nuevaSolicitud)
    
    If idGenerado > 0 Then
        nuevaSolicitud.idSolicitud = idGenerado
        
        ' Logging de auditoría: Éxito
        If Not m_OperationLogger Is Nothing Then
            m_OperationLogger.LogOperation "CREATE_SOLICITUD_SUCCESS", "Solicitud creada con ID: " & idGenerado, "CSolicitudService.CreateSolicitud"
        End If
        
        Set CreateSolicitud = nuevaSolicitud
    Else
        ' Logging de auditoría: Fallo
        If Not m_OperationLogger Is Nothing Then
            m_OperationLogger.LogOperation "CREATE_SOLICITUD_FAILED", "Error al persistir solicitud", "CSolicitudService.CreateSolicitud"
        End If
        Set CreateSolicitud = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "CREATE_SOLICITUD_ERROR", "Error: " & Err.Description, "CSolicitudService.CreateSolicitud"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.CreateSolicitud"
    Set CreateSolicitud = Nothing
End Function

' Obtiene una solicitud por su ID
Public Function GetSolicitud(ByVal idSolicitud As Long) As T_Solicitud
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "GET_SOLICITUD", "Obteniendo solicitud ID: " & idSolicitud, "CSolicitudService.GetSolicitud"
    End If
    
    ' Delegar directamente al repositorio
    Set GetSolicitud = m_SolicitudRepository.ObtenerSolicitudPorId(idSolicitud)
    
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "GET_SOLICITUD_ERROR", "Error: " & Err.Description, "CSolicitudService.GetSolicitud"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.GetSolicitud"
    Set GetSolicitud = Nothing
End Function

' Guarda una solicitud existente
Public Function SaveSolicitud(solicitud As T_Solicitud) As Boolean
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "SAVE_SOLICITUD", "Guardando solicitud ID: " & solicitud.idSolicitud, "CSolicitudService.SaveSolicitud"
    End If
    
    ' Delegar al repositorio y evaluar resultado
    Dim idResultado As Long
    idResultado = m_SolicitudRepository.GuardarSolicitud(solicitud)
    
    SaveSolicitud = (idResultado > 0)
    
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "SAVE_SOLICITUD_ERROR", "Error: " & Err.Description, "CSolicitudService.SaveSolicitud"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.SaveSolicitud"
    SaveSolicitud = False
End Function

' Cambia el estado de una solicitud con validación de transición
Public Function ChangeState(ByVal idSolicitud As Long, ByVal newState As String) As Boolean
    On Error GoTo ErrorHandler
    ' TODO: Implementar lógica de validación de transición aquí

    Dim solicitud As T_Solicitud
    Set solicitud = m_SolicitudRepository.ObtenerSolicitudPorId(idSolicitud)

    If solicitud Is Nothing Then
        ' La solicitud no existe
        ChangeState = False
        Exit Function
    End If

    ' Validar si la transición del estado actual al nuevo es válida
    ' If Not IsValidTransition(solicitud.estadoInterno, newState) Then
    '     ChangeState = False
    '     Exit Function
    ' End If

    solicitud.estadoInterno = newState
    ChangeState = Me.SaveSolicitud(solicitud)

    If ChangeState Then
        m_OperationLogger.LogOperation "CHANGE_STATE", "Estado de Solicitud " & idSolicitud & " cambiado a " & newState, "CSolicitudService"
    End If

    Exit Function
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.ChangeState"
    ChangeState = False
End Function

' ============================================================================
' IMPLEMENTACIONES DE LA INTERFAZ ISolicitudService
' ============================================================================

Private Function ISolicitudService_CreateSolicitud(ByVal idExpediente As String, ByVal tipo As String, ByVal subTipo As String) As T_Solicitud
    Set ISolicitudService_CreateSolicitud = CreateSolicitud(idExpediente, tipo, subTipo)
End Function

Private Function ISolicitudService_GetSolicitud(ByVal idSolicitud As Long) As T_Solicitud
    Set ISolicitudService_GetSolicitud = GetSolicitud(idSolicitud)
End Function

Private Function ISolicitudService_SaveSolicitud(solicitud As T_Solicitud) As Boolean
    ISolicitudService_SaveSolicitud = SaveSolicitud(solicitud)
End Function

Private Function ISolicitudService_ChangeState(ByVal idSolicitud As Long, ByVal newState As String) As Boolean
    ISolicitudService_ChangeState = ChangeState(idSolicitud, newState)
End Function











