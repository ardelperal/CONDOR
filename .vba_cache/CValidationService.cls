Option Compare Database
Option Explicit
' ===================================================================
' Clase: CValidationService
' Descripción: Implementación concreta del servicio de validaciones de negocio
' Autor: Sistema CONDOR
' Fecha: 2024
' ===================================================================


' Implementa la interfaz IValidationService
Implements IValidationService

Private m_OperationLogger As IOperationLogger

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN PRINCIPAL
' ===================================================================

Private Function IValidationService_ValidarSolicitud(ByVal solicitud As T_Solicitud, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría: Inicio de validación
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "VALIDATION_START", "Iniciando validación solicitud ID: " & solicitud.idSolicitud & ", Tipo: " & solicitud.tipoSolicitud, "CValidationService.ValidarSolicitud"
    End If
    
    MensajeError = ""
    
    ' Validar campos obligatorios
    If Not IValidationService_ValidarCamposObligatorios(solicitud, MensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar número de expediente
    If Not IValidationService_ValidarNumeroExpediente(solicitud.NumeroExpediente, MensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar tipo de solicitud
    If Not IValidationService_ValidarTipoSolicitud(solicitud.tipoSolicitud, MensajeError) Then
        IValidationService_ValidarSolicitud = False
        Exit Function
    End If
    
    ' Validar importes (solo para solicitudes PC)
    If solicitud.tipoSolicitud = "PC" Then
        If Not IValidationService_ValidarImportes(solicitud.importeOriginal, solicitud.importeNuevo, MensajeError) Then
            IValidationService_ValidarSolicitud = False
            Exit Function
        End If
    End If
    
    ' Si llegamos aquí, todas las validaciones pasaron
    ' Logging de auditoría: Validación exitosa
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "VALIDATION_SUCCESS", "Validación exitosa para solicitud ID: " & solicitud.idSolicitud, "CValidationService.ValidarSolicitud"
    End If
    
    IValidationService_ValidarSolicitud = True
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "VALIDATION_ERROR", "Error en validación solicitud ID: " & solicitud.idSolicitud & " - " & Err.Description, "CValidationService.ValidarSolicitud"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarSolicitud"
    MensajeError = "Error interno en validación: " & Err.Description
    IValidationService_ValidarSolicitud = False
End Function

' ===================================================================
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN ESPECÍFICA
' ===================================================================

Private Function IValidationService_ValidarNumeroExpediente(ByVal NumeroExpediente As String, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que no esté vacío
    If Trim(NumeroExpediente) = "" Then
        MensajeError = "El número de expediente es obligatorio"
        IValidationService_ValidarNumeroExpediente = False
        Exit Function
    End If
    
    ' Validar formato básico (debe contener al menos 3 caracteres)
    If Len(Trim(NumeroExpediente)) < 3 Then
        MensajeError = "El número de expediente debe tener al menos 3 caracteres"
        IValidationService_ValidarNumeroExpediente = False
        Exit Function
    End If
    
    ' Validación exitosa
    IValidationService_ValidarNumeroExpediente = True
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarNumeroExpediente"
    MensajeError = "Error validando número de expediente: " & Err.Description
    IValidationService_ValidarNumeroExpediente = False
End Function

Private Function IValidationService_ValidarTipoSolicitud(ByVal tipoSolicitud As String, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que no esté vacío
    If Trim(tipoSolicitud) = "" Then
        MensajeError = "El tipo de solicitud es obligatorio"
        IValidationService_ValidarTipoSolicitud = False
        Exit Function
    End If
    
    ' Validar que sea un tipo válido
    Select Case UCase(Trim(tipoSolicitud))
        Case "PC", "CD_CA", "CD_CA_SUB"
            IValidationService_ValidarTipoSolicitud = True
        Case Else
            MensajeError = "Tipo de solicitud no válido. Debe ser PC, CD_CA o CD_CA_SUB"
            IValidationService_ValidarTipoSolicitud = False
    End Select
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarTipoSolicitud"
    MensajeError = "Error validando tipo de solicitud: " & Err.Description
    IValidationService_ValidarTipoSolicitud = False
End Function

Private Function IValidationService_ValidarCamposObligatorios(ByVal solicitud As T_Solicitud, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar descripción
    If Trim(solicitud.Descripcion) = "" Then
        MensajeError = "La descripción es obligatoria"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar justificación
    If Trim(solicitud.justificacionCambio) = "" Then
        MensajeError = "La justificación es obligatoria"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar usuario creador
    If Trim(solicitud.UsuarioCreador) = "" Then
        MensajeError = "El usuario creador es obligatorio"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Validar estado
    If Trim(solicitud.Estado) = "" Then
        MensajeError = "El estado es obligatorio"
        IValidationService_ValidarCamposObligatorios = False
        Exit Function
    End If
    
    ' Todos los campos obligatorios están presentes
    IValidationService_ValidarCamposObligatorios = True
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarCamposObligatorios"
    MensajeError = "Error validando campos obligatorios: " & Err.Description
    IValidationService_ValidarCamposObligatorios = False
End Function

Private Function IValidationService_ValidarImportes(ByVal importeOriginal As Currency, ByVal importeNuevo As Currency, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que los importes sean positivos
    If importeOriginal <= 0 Then
        MensajeError = "El importe original debe ser mayor que cero"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    If importeNuevo <= 0 Then
        MensajeError = "El importe nuevo debe ser mayor que cero"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    ' Validar que haya una diferencia significativa (al menos 1%)
    Dim porcentajeCambio As Double
    porcentajeCambio = Abs((importeNuevo - importeOriginal) / importeOriginal) * 100
    
    If porcentajeCambio < 1 Then
        MensajeError = "La diferencia entre importes debe ser al menos del 1%"
        IValidationService_ValidarImportes = False
        Exit Function
    End If
    
    ' Validación exitosa
    IValidationService_ValidarImportes = True
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarImportes"
    MensajeError = "Error validando importes: " & Err.Description
    IValidationService_ValidarImportes = False
End Function

Private Function IValidationService_ValidarTransicionEstado(ByVal estadoActual As String, ByVal estadoNuevo As String, ByRef MensajeError As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría: Validación de transición de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "VALIDATION_TRANSITION", "Validando transición de estado - Actual: " & estadoActual & ", Nuevo: " & estadoNuevo, "CValidationService.ValidarTransicionEstado"
    End If
    
    ' Validar que los estados no estén vacíos
    If Trim(estadoActual) = "" Or Trim(estadoNuevo) = "" Then
        MensajeError = "Los estados actual y nuevo son obligatorios"
        IValidationService_ValidarTransicionEstado = False
        Exit Function
    End If
    
    ' Validar transiciones permitidas según las reglas de negocio
    Select Case UCase(Trim(estadoActual))
        Case "BORRADOR"
            ' Desde Borrador se puede ir a Enviado o Cancelado
            If UCase(Trim(estadoNuevo)) = "ENVIADO" Or UCase(Trim(estadoNuevo)) = "CANCELADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                MensajeError = "Desde Borrador solo se puede pasar a Enviado o Cancelado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "ENVIADO"
            ' Desde Enviado se puede ir a Aprobado, Rechazado o Cancelado
            If UCase(Trim(estadoNuevo)) = "APROBADO" Or UCase(Trim(estadoNuevo)) = "RECHAZADO" Or UCase(Trim(estadoNuevo)) = "CANCELADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                MensajeError = "Desde Enviado solo se puede pasar a Aprobado, Rechazado o Cancelado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "APROBADO"
            ' Desde Aprobado se puede ir a Finalizado
            If UCase(Trim(estadoNuevo)) = "FINALIZADO" Then
                IValidationService_ValidarTransicionEstado = True
            Else
                MensajeError = "Desde Aprobado solo se puede pasar a Finalizado"
                IValidationService_ValidarTransicionEstado = False
            End If
            
        Case "RECHAZADO", "CANCELADO", "FINALIZADO"
            ' Estados finales - no se puede cambiar
            MensajeError = "No se puede cambiar el estado desde " & estadoActual
            IValidationService_ValidarTransicionEstado = False
            
        Case Else
            MensajeError = "Estado actual no reconocido: " & estadoActual
            IValidationService_ValidarTransicionEstado = False
    End Select
    
    ' Logging de auditoría: Validación de transición exitosa
    If Not m_OperationLogger Is Nothing Then
        If IValidationService_ValidarTransicionEstado Then
            m_OperationLogger.LogOperation "VALIDATION_TRANSITION_SUCCESS", "Transición válida - Actual: " & estadoActual & ", Nuevo: " & estadoNuevo, "CValidationService.ValidarTransicionEstado"
        Else
            m_OperationLogger.LogOperation "VALIDATION_TRANSITION_FAILED", "Transición inválida - Actual: " & estadoActual & ", Nuevo: " & estadoNuevo & ", Error: " & MensajeError, "CValidationService.ValidarTransicionEstado"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "VALIDATION_ERROR", "Error validando transición - Actual: " & estadoActual & ", Nuevo: " & estadoNuevo & ", Error: " & Err.Description, "CValidationService.ValidarTransicionEstado"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.ValidarTransicionEstado"
    MensajeError = "Error validando transición de estado: " & Err.Description
    IValidationService_ValidarTransicionEstado = False
End Function

' Método de inicialización para inyectar dependencias
Public Sub Initialize(operationLogger As IOperationLogger)
    On Error GoTo ErrorHandler
    
    Set m_OperationLogger = operationLogger
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CValidationService.Initialize"
End Sub






