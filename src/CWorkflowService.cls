Option Compare Database
Option Explicit
'******************************************************************************
' CLASE: CWorkflowService
' DESCRIPCIÓN: Implementación concreta del servicio de workflow y estados
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Lección 1 - Implementa la interfaz IWorkflowService
'        Lección 3 - Sigue la arquitectura de 3 capas (Capa de Servicios)
'        Lección 6 - Utiliza módulos de acceso a datos como especificación
'******************************************************************************


' Lección 1: Implementar la interfaz
Implements IWorkflowService

'******************************************************************************
' CONSTANTES PRIVADAS
'******************************************************************************

Private Const ESTADO_BORRADOR As String = "Borrador"
Private Const ESTADO_EN_PROCESO As String = "EnProceso"
Private Const ESTADO_APROBADO As String = "Aprobado"
Private Const ESTADO_RECHAZADO As String = "Rechazado"
Private Const ESTADO_CANCELADO As String = "Cancelado"

Private Const ROL_USUARIO As String = "Usuario"
Private Const ROL_APROBADOR As String = "Aprobador"
Private Const ROL_ADMINISTRADOR As String = "Administrador"

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN DE TRANSICIONES
'******************************************************************************

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, _
                                                    ByVal estadoOrigen As String, _
                                                    ByVal estadoDestino As String, _
                                                    ByVal tipoSolicitud As String, _
                                                    ByVal usuarioRol As String) As Boolean
    
    ' Logging de auditoría: Validación de transición de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_TRANSITION_VALIDATION", "Validando transición solicitud ID: " & SolicitudID & ", de " & estadoOrigen & " a " & estadoDestino & " por usuario rol: " & usuarioRol, "CWorkflowService.ValidateTransition"
    End If
    
    Dim resultado As Boolean
    resultado = False
    
    ' Validar parámetros de entrada
    If Len(Trim(estadoOrigen)) = 0 Or Len(Trim(estadoDestino)) = 0 Or _
       Len(Trim(tipoSolicitud)) = 0 Or Len(Trim(usuarioRol)) = 0 Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Verificar si la transición existe en TbTransiciones
    If Not TransitionExists(estadoOrigen, estadoDestino, tipoSolicitud) Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Verificar permisos del usuario para esta transición
    If Not IWorkflowService_HasTransitionPermission(usuarioRol, estadoOrigen, estadoDestino, tipoSolicitud) Then
        IWorkflowService_ValidateTransition = False
        Exit Function
    End If
    
    ' Si llegamos aquí, la transición es válida
    IWorkflowService_ValidateTransition = True
    
    ' Logging de auditoría: Transición validada exitosamente
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_TRANSITION_VALIDATED", "Transición validada exitosamente para solicitud ID: " & SolicitudID & ", de " & estadoOrigen & " a " & estadoDestino, "CWorkflowService.ValidateTransition"
    End If
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE OBTENCIÓN DE ESTADOS
'******************************************************************************

Private Function IWorkflowService_GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Dim estados As New Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Consultar estados activos para el tipo de solicitud
    sql = "SELECT CodigoEstado, NombreEstado FROM TbEstados " & _
          "WHERE TipoSolicitud = '" & tipoSolicitud & "' AND Activo = True " & _
          "ORDER BY NombreEstado"
    
    Set rs = db.OpenRecordset(sql)
    
    Do While Not rs.EOF
        estados.Add rs("CodigoEstado").value, rs("CodigoEstado").value
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Set IWorkflowService_GetAvailableStates = estados
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, _
                                               ByVal tipoSolicitud As String, _
                                               ByVal usuarioRol As String) As Collection
    
    Dim estadosSiguientes As New Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Consultar transiciones válidas desde el estado actual
    sql = "SELECT DISTINCT t.EstadoDestinoID, e.CodigoEstado, e.NombreEstado " & _
          "FROM TbTransiciones t INNER JOIN TbEstados e ON t.EstadoDestinoID = e.ID " & _
          "INNER JOIN TbEstados eo ON t.EstadoOrigenID = eo.ID " & _
          "WHERE eo.CodigoEstado = '" & estadoActual & "' " & _
          "AND t.TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND t.Activo = True AND e.Activo = True " & _
          "AND (t.RolRequerido IS NULL OR t.RolRequerido = '' OR t.RolRequerido = '" & usuarioRol & "') " & _
          "ORDER BY e.NombreEstado"
    
    Set rs = db.OpenRecordset(sql)
    
    Do While Not rs.EOF
        estadosSiguientes.Add rs("CodigoEstado").value, rs("CodigoEstado").value
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Set IWorkflowService_GetNextStates = estadosSiguientes
End Function

Private Function IWorkflowService_GetInitialState(ByVal tipoSolicitud As String) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As String
    
    Set db = CurrentDb
    
    sql = "SELECT CodigoEstado FROM TbEstados " & _
          "WHERE TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND EsEstadoInicial = True AND Activo = True"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        resultado = rs("CodigoEstado").value
    Else
        resultado = ESTADO_BORRADOR ' Estado por defecto
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    IWorkflowService_GetInitialState = resultado
End Function

Private Function IWorkflowService_IsStateFinal(ByVal estadoCodigo As String, _
                                              ByVal tipoSolicitud As String) As Boolean
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As Boolean
    
    Set db = CurrentDb
    
    sql = "SELECT EsEstadoFinal FROM TbEstados " & _
          "WHERE CodigoEstado = '" & estadoCodigo & "' " & _
          "AND TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND Activo = True"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        resultado = CBool(rs("EsEstadoFinal").value)
    Else
        resultado = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    IWorkflowService_IsStateFinal = resultado
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE HISTORIAL DE ESTADOS
'******************************************************************************

Private Function IWorkflowService_RecordStateChange(ByVal SolicitudID As Long, _
                                                   ByVal estadoAnterior As String, _
                                                   ByVal estadoNuevo As String, _
                                                   ByVal usuario As String, _
                                                   ByVal comentarios As String) As Boolean
    
    ' Logging de auditoría: Registro de cambio de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_STATE_CHANGE", "Registrando cambio de estado para solicitud ID: " & SolicitudID & ", de " & estadoAnterior & " a " & estadoNuevo & " por usuario: " & usuario, "CWorkflowService.RecordStateChange"
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    
    ' Insertar registro en TbHistorialEstados (asumiendo que existe esta tabla)
    sql = "INSERT INTO TbHistorialEstados (SolicitudID, EstadoAnterior, EstadoNuevo, Usuario, Comentarios, FechaCambio) " & _
          "VALUES (" & SolicitudID & ", '" & estadoAnterior & "', '" & estadoNuevo & "', '" & usuario & "', '" & comentarios & "', Now())"
    
    db.Execute sql
    
    Set db = Nothing
    
    ' Logging de auditoría: Cambio de estado registrado exitosamente
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_STATE_RECORDED", "Cambio de estado registrado exitosamente para solicitud ID: " & SolicitudID & ", de " & estadoAnterior & " a " & estadoNuevo, "CWorkflowService.RecordStateChange"
    End If
    
    IWorkflowService_RecordStateChange = True
    Exit Function
    
ErrorHandler:
    ' Logging de auditoría: Error al registrar cambio de estado
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "WORKFLOW_STATE_RECORD_ERROR", "Error al registrar cambio de estado para solicitud ID: " & SolicitudID & ". Error: " & Err.Description, "CWorkflowService.RecordStateChange"
    End If
    
    modErrorHandler.LogError Err.Number, Err.Description, "CWorkflowService.RecordStateChange"
    If Not db Is Nothing Then Set db = Nothing
    IWorkflowService_RecordStateChange = False
End Function

Private Function IWorkflowService_GetStateHistory(ByVal SolicitudID As Long) As Collection
    On Error GoTo ErrorHandler
    
    Dim historial As New Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim registro As String
    
    Set db = CurrentDb
    
    sql = "SELECT EstadoAnterior, EstadoNuevo, Usuario, Comentarios, FechaCambio " & _
          "FROM TbHistorialEstados " & _
          "WHERE SolicitudID = " & SolicitudID & " " & _
          "ORDER BY FechaCambio DESC"
    
    Set rs = db.OpenRecordset(sql)
    
    Do While Not rs.EOF
        registro = rs("FechaCambio").value & " - " & rs("Usuario").value & ": " & _
                  rs("EstadoAnterior").value & " ? " & rs("EstadoNuevo").value
        If Not IsNull(rs("Comentarios").value) And Len(Trim(rs("Comentarios").value)) > 0 Then
            registro = registro & " (" & rs("Comentarios").value & ")"
        End If
        historial.Add registro
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Set IWorkflowService_GetStateHistory = historial
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CWorkflowService.GetStateHistory"
    If Not rs Is Nothing Then Set rs = Nothing
    If Not db Is Nothing Then Set db = Nothing
    Set IWorkflowService_GetStateHistory = New Collection
End Function

'******************************************************************************
' IMPLEMENTACIÓN DE MÉTODOS DE VALIDACIÓN DE PERMISOS
'******************************************************************************

Private Function IWorkflowService_HasTransitionPermission(ByVal usuarioRol As String, _
                                                         ByVal estadoOrigen As String, _
                                                         ByVal estadoDestino As String, _
                                                         ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As Boolean
    
    Set db = CurrentDb
    
    ' Consultar si existe la transición y si el rol tiene permisos
    sql = "SELECT t.RolRequerido FROM TbTransiciones t " & _
          "INNER JOIN TbEstados eo ON t.EstadoOrigenID = eo.ID " & _
          "INNER JOIN TbEstados ed ON t.EstadoDestinoID = ed.ID " & _
          "WHERE eo.CodigoEstado = '" & estadoOrigen & "' " & _
          "AND ed.CodigoEstado = '" & estadoDestino & "' " & _
          "AND t.TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND t.Activo = True"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        Dim rolRequerido As String
        rolRequerido = Nz(rs("RolRequerido").value, "")
        
        ' Si no hay rol requerido específico, cualquier usuario puede hacer la transición
        If Len(Trim(rolRequerido)) = 0 Then
            resultado = True
        ' Si el usuario es administrador, puede hacer cualquier transición
        ElseIf usuarioRol = ROL_ADMINISTRADOR Then
            resultado = True
        ' Verificar si el rol del usuario coincide con el requerido
        Else
            resultado = (usuarioRol = rolRequerido)
        End If
    Else
        resultado = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    IWorkflowService_HasTransitionPermission = resultado
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CWorkflowService.HasTransitionPermission"
    If Not rs Is Nothing Then Set rs = Nothing
    If Not db Is Nothing Then Set db = Nothing
    IWorkflowService_HasTransitionPermission = False
End Function

Private Function IWorkflowService_RequiresApproval(ByVal estadoOrigen As String, _
                                                  ByVal estadoDestino As String, _
                                                  ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As Boolean
    
    Set db = CurrentDb
    
    sql = "SELECT t.RequiereAprobacion FROM TbTransiciones t " & _
          "INNER JOIN TbEstados eo ON t.EstadoOrigenID = eo.ID " & _
          "INNER JOIN TbEstados ed ON t.EstadoDestinoID = ed.ID " & _
          "WHERE eo.CodigoEstado = '" & estadoOrigen & "' " & _
          "AND ed.CodigoEstado = '" & estadoDestino & "' " & _
          "AND t.TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND t.Activo = True"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        resultado = CBool(rs("RequiereAprobacion").value)
    Else
        resultado = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    IWorkflowService_RequiresApproval = resultado
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CWorkflowService.RequiresApproval"
    If Not rs Is Nothing Then Set rs = Nothing
    If Not db Is Nothing Then Set db = Nothing
    IWorkflowService_RequiresApproval = False
End Function

'******************************************************************************
' MÉTODOS AUXILIARES PRIVADOS
'******************************************************************************

Private Function TransitionExists(ByVal estadoOrigen As String, _
                                 ByVal estadoDestino As String, _
                                 ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim resultado As Boolean
    
    Set db = CurrentDb
    
    sql = "SELECT COUNT(*) AS Total FROM TbTransiciones t " & _
          "INNER JOIN TbEstados eo ON t.EstadoOrigenID = eo.ID " & _
          "INNER JOIN TbEstados ed ON t.EstadoDestinoID = ed.ID " & _
          "WHERE eo.CodigoEstado = '" & estadoOrigen & "' " & _
          "AND ed.CodigoEstado = '" & estadoDestino & "' " & _
          "AND t.TipoSolicitud = '" & tipoSolicitud & "' " & _
          "AND t.Activo = True"
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        resultado = (rs("Total").value > 0)
    Else
        resultado = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    TransitionExists = resultado
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CWorkflowService.TransitionExists"
    If Not rs Is Nothing Then Set rs = Nothing
    If Not db Is Nothing Then Set db = Nothing
    TransitionExists = False
End Function













