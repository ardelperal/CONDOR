VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'==============================================================================
' Clase: CWorkflowRepository
' Propósito: Implementación real del repositorio de workflow que accede a la base de datos
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================


Implements IWorkflowRepository

'==============================================================================
' VARIABLES PRIVADAS
'==============================================================================

' Dependencia para acceso a configuración
Private m_Config As IConfig

'==============================================================================
' MÉTODOS PÚBLICOS
'==============================================================================

'''
' Inicializa el repositorio con sus dependencias
' @param config: Servicio de configuración para obtener rutas de base de datos
'''
Public Sub Initialize(config As IConfig)
    Set m_Config = config
End Sub

'==============================================================================
' IMPLEMENTACIÓN DE INTERFAZ IWorkflowRepository
'==============================================================================

'''
' Valida si una transición de estado es permitida consultando la base de datos
' REFACTORIZADO: Usa consultas parametrizadas para prevenir inyección SQL
' @param tipoSolicitud: Tipo de solicitud (PC, CDCA, CDCASUB)
' @param estadoOrigen: Estado actual de la solicitud
' @param estadoDestino: Estado al que se quiere transicionar
' @return Boolean: True si la transición es válida, False en caso contrario
'''
Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    Dim ErrorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim validCount As Long
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada (SEGURA)
    Set qdf = db.CreateQueryDef("", _
        "SELECT COUNT(*) AS Valid " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND Origen.CodigoEstado = [pEstadoOrigen] " & _
        "AND Destino.CodigoEstado = [pEstadoDestino] " & _
        "AND TbTransiciones.Activo = True;")
    
    ' Asignar valores a los parámetros de forma segura
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    qdf.parameters("pEstadoOrigen").value = estadoOrigen
    qdf.parameters("pEstadoDestino").value = estadoDestino
    
    ' Ejecutar la consulta parametrizada
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Verificar si hay resultados
    If Not rs.EOF Then
        validCount = rs.Fields("Valid").value
        IWorkflowRepository_IsValidTransition = (validCount > 0)
    Else
        IWorkflowRepository_IsValidTransition = False
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir que la transición no es válida
    IWorkflowRepository_IsValidTransition = False
    
    ' Limpiar recursos si está abiertos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    ' Registrar el error para debugging
    Debug.Print "Error en CWorkflowRepository.IsValidTransition: " & Err.Number & " - " & Err.Description
End Function

'''
' Obtiene los estados disponibles para un tipo de solicitud específico
' @param tipoSolicitud: Tipo de solicitud (PC, CDCA, CDCASUB)
' @return Collection: Collection con los códigos de estados disponibles
'''
Private Function IWorkflowRepository_GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim colEstados As New Collection
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT DISTINCT E.CodigoEstado, E.NombreEstado " & _
        "FROM TbEstados E " & _
        "INNER JOIN TbTransiciones T ON E.ID = T.EstadoOrigenID OR E.ID = T.EstadoDestinoID " & _
        "WHERE T.TipoSolicitud = [pTipoSolicitud] AND T.Activo = True " & _
        "ORDER BY E.NombreEstado;")
    
    ' Asignar parámetro
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    
    ' Ejecutar consulta y obtener recordset
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Recorrer el recordset y poblar la colección
    Do While Not rs.EOF
        colEstados.Add rs.Fields("CodigoEstado").value
        rs.MoveNext
    Loop
    
    ' Cerrar recordset y conexión
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver la colección poblada
    Set IWorkflowRepository_GetAvailableStates = colEstados
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    Set IWorkflowRepository_GetAvailableStates = Nothing
    Debug.Print "Error en CWorkflowRepository.GetAvailableStates: " & Err.Number & " - " & Err.Description
End Function

'''
' Obtiene los estados siguientes válidos desde un estado actual
' @param estadoActual: Estado actual de la solicitud
' @param tipoSolicitud: Tipo de solicitud
' @param usuarioRol: Rol del usuario actual
' @return Collection: Collection con los códigos de estados siguientes válidos
'''
Private Function IWorkflowRepository_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim colEstados As New Collection
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT DISTINCT Destino.CodigoEstado, Destino.NombreEstado " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE Origen.CodigoEstado = [pEstadoActual] " & _
        "AND TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND (TbTransiciones.RolRequerido IS NULL OR TbTransiciones.RolRequerido = [pUsuarioRol]) " & _
        "AND TbTransiciones.Activo = True " & _
        "ORDER BY Destino.NombreEstado;")
    
    ' Asignar parámetros
    qdf.parameters("pEstadoActual").value = estadoActual
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    qdf.parameters("pUsuarioRol").value = usuarioRol
    
    ' Ejecutar consulta y obtener recordset
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Recorrer el recordset y poblar la colección
    Do While Not rs.EOF
        colEstados.Add rs.Fields("CodigoEstado").value
        rs.MoveNext
    Loop
    
    ' Cerrar recordset y conexión
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver la colección poblada
    Set IWorkflowRepository_GetNextStates = colEstados
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    Set IWorkflowRepository_GetNextStates = Nothing
    Debug.Print "Error en CWorkflowRepository.GetNextStates: " & Err.Number & " - " & Err.Description
End Function

'''
' Obtiene el estado inicial para un tipo de solicitud específico
' @param tipoSolicitud: Tipo de solicitud (PC, CDCA, CDCASUB)
' @return String: Código del estado inicial
'''
Private Function IWorkflowRepository_GetInitialState(ByVal tipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT E.CodigoEstado " & _
        "FROM TbEstados E " & _
        "INNER JOIN TbTiposSolicitud TS ON E.ID = TS.EstadoInicialID " & _
        "WHERE TS.Codigo = [pTipoSolicitud];")
    
    ' Asignar parámetro
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Obtener resultado
    If Not rs.EOF Then
        IWorkflowRepository_GetInitialState = rs.Fields("CodigoEstado").value
    Else
        IWorkflowRepository_GetInitialState = ""
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    IWorkflowRepository_GetInitialState = ""
    Debug.Print "Error en CWorkflowRepository.GetInitialState: " & Err.Number & " - " & Err.Description
End Function

'''
' Verifica si un estado es final (no tiene transiciones salientes)
' @param estadoCodigo: Código del estado a verificar
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si el estado es final
'''
Private Function IWorkflowRepository_IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim transitionCount As Long
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT COUNT(*) AS TransitionCount " & _
        "FROM TbTransiciones T " & _
        "INNER JOIN TbEstados E ON T.EstadoOrigenID = E.ID " & _
        "WHERE E.CodigoEstado = [pEstadoCodigo] " & _
        "AND T.TipoSolicitud = [pTipoSolicitud] " & _
        "AND T.Activo = True;")
    
    ' Asignar parámetros
    qdf.parameters("pEstadoCodigo").value = estadoCodigo
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Verificar si hay transiciones salientes
    If Not rs.EOF Then
        transitionCount = rs.Fields("TransitionCount").value
        IWorkflowRepository_IsStateFinal = (transitionCount = 0)
    Else
        IWorkflowRepository_IsStateFinal = True
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    IWorkflowRepository_IsStateFinal = True
    Debug.Print "Error en CWorkflowRepository.IsStateFinal: " & Err.Number & " - " & Err.Description
End Function

'''
' Registra un cambio de estado en el historial
' @param SolicitudID: ID de la solicitud
' @param estadoAnterior: Estado anterior
' @param estadoNuevo: Nuevo estado
' @param usuario: Usuario que realizó el cambio
' @param comentarios: Comentarios del cambio
' @return Boolean: True si el registro fue exitoso
'''
Private Function IWorkflowRepository_RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada INSERT
    Set qdf = db.CreateQueryDef("", _
        "INSERT INTO TbHistorialEstados (SolicitudID, EstadoAnterior, EstadoNuevo, Usuario, FechaCambio, Comentarios) " & _
        "VALUES ([pSolicitudID], [pEstadoAnterior], [pEstadoNuevo], [pUsuario], Now(), [pComentarios]);")
    
    ' Asignar parámetros
    qdf.parameters("pSolicitudID").value = SolicitudID
    qdf.parameters("pEstadoAnterior").value = estadoAnterior
    qdf.parameters("pEstadoNuevo").value = estadoNuevo
    qdf.parameters("pUsuario").value = usuario
    qdf.parameters("pComentarios").value = comentarios
    
    ' Ejecutar la inserción
    qdf.Execute dbFailOnError
    
    ' Limpiar recursos
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Retornar éxito
    IWorkflowRepository_RecordStateChange = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en CWorkflowRepository.RecordStateChange: " & Err.Number & " - " & Err.Description
    
    ' Limpiar recursos en caso de error
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    ' Retornar fallo
    IWorkflowRepository_RecordStateChange = False
End Function

'''
' Obtiene el historial de cambios de estado para una solicitud
' @param SolicitudID: ID de la solicitud
' @return Collection: Collection con el historial de cambios
'''
Private Function IWorkflowRepository_GetStateHistory(ByVal SolicitudID As Long) As Collection
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim colHistorial As New Collection
    Dim historyItem As String
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT EstadoAnterior, EstadoNuevo, Usuario, FechaCambio, Comentarios " & _
        "FROM TbHistorialEstados " & _
        "WHERE SolicitudID = [pSolicitudID] " & _
        "ORDER BY FechaCambio DESC;")
    
    ' Asignar parámetro
    qdf.parameters("pSolicitudID").value = SolicitudID
    
    ' Ejecutar consulta y obtener recordset
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Recorrer el recordset y poblar la colección
    Do While Not rs.EOF
        ' Crear un string con la información del historial
        historyItem = rs.Fields("EstadoAnterior").value & "|" & _
                     rs.Fields("EstadoNuevo").value & "|" & _
                     rs.Fields("Usuario").value & "|" & _
                     rs.Fields("FechaCambio").value & "|" & _
                     Nz(rs.Fields("Comentarios").value, "")
        colHistorial.Add historyItem
        rs.MoveNext
    Loop
    
    ' Cerrar recordset y conexión
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver la colección poblada
    Set IWorkflowRepository_GetStateHistory = colHistorial
    
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not qdf Is Nothing Then
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    Set IWorkflowRepository_GetStateHistory = Nothing
    Debug.Print "Error en CWorkflowRepository.GetStateHistory: " & Err.Number & " - " & Err.Description
End Function

'''
' Verifica si un usuario tiene permisos para realizar una transición específica
' @param usuarioRol: Rol del usuario
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si tiene permisos
'''
Private Function IWorkflowRepository_HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim permissionCount As Long
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT COUNT(*) AS PermissionCount " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE Origen.CodigoEstado = [pEstadoOrigen] " & _
        "AND Destino.CodigoEstado = [pEstadoDestino] " & _
        "AND TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND (TbTransiciones.RolRequerido IS NULL OR TbTransiciones.RolRequerido = [pUsuarioRol]) " & _
        "AND TbTransiciones.Activo = True;")
    
    ' Asignar parámetros
    qdf.parameters("pEstadoOrigen").value = estadoOrigen
    qdf.parameters("pEstadoDestino").value = estadoDestino
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    qdf.parameters("pUsuarioRol").value = usuarioRol
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Verificar permisos
    If Not rs.EOF Then
        permissionCount = rs.Fields("PermissionCount").value
        IWorkflowRepository_HasTransitionPermission = (permissionCount > 0)
    Else
        IWorkflowRepository_HasTransitionPermission = False
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    IWorkflowRepository_HasTransitionPermission = False
    Debug.Print "Error en CWorkflowRepository.HasTransitionPermission: " & Err.Number & " - " & Err.Description
End Function

'''
' Verifica si una transición requiere aprobación adicional
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return Boolean: True si requiere aprobación
'''
Private Function IWorkflowRepository_RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT TbTransiciones.RequiereAprobacion " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE Origen.CodigoEstado = [pEstadoOrigen] " & _
        "AND Destino.CodigoEstado = [pEstadoDestino] " & _
        "AND TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND TbTransiciones.Activo = True;")
    
    ' Asignar parámetros
    qdf.parameters("pEstadoOrigen").value = estadoOrigen
    qdf.parameters("pEstadoDestino").value = estadoDestino
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Verificar si requiere aprobación
    If Not rs.EOF Then
        IWorkflowRepository_RequiresApproval = CBool(rs.Fields("RequiereAprobacion").value)
    Else
        IWorkflowRepository_RequiresApproval = False
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    IWorkflowRepository_RequiresApproval = False
    Debug.Print "Error en CWorkflowRepository.RequiresApproval: " & Err.Number & " - " & Err.Description
End Function

'''
' Obtiene el rol requerido para realizar una transición específica
' @param estadoOrigen: Estado origen de la transición
' @param estadoDestino: Estado destino de la transición
' @param tipoSolicitud: Tipo de solicitud
' @return String: Rol requerido para la transición (vacío si no se requiere rol específico)
'''
Private Function IWorkflowRepository_GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    ' Obtener ruta del backend desde configuración
    Dim backendPath As String
    backendPath = m_Config.GetDataPath()
    
    ' Abrir conexión a la base de datos backend
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False)
    
    ' Crear QueryDef con consulta parametrizada
    Set qdf = db.CreateQueryDef("", _
        "SELECT TbTransiciones.RolRequerido " & _
        "FROM (TbTransiciones " & _
        "INNER JOIN TbEstados AS Origen ON TbTransiciones.EstadoOrigenID = Origen.ID) " & _
        "INNER JOIN TbEstados AS Destino ON TbTransiciones.EstadoDestinoID = Destino.ID " & _
        "WHERE Origen.CodigoEstado = [pEstadoOrigen] " & _
        "AND Destino.CodigoEstado = [pEstadoDestino] " & _
        "AND TbTransiciones.TipoSolicitud = [pTipoSolicitud] " & _
        "AND TbTransiciones.Activo = True;")
    
    ' Asignar parámetros
    qdf.parameters("pEstadoOrigen").value = estadoOrigen
    qdf.parameters("pEstadoDestino").value = estadoDestino
    qdf.parameters("pTipoSolicitud").value = tipoSolicitud
    
    ' Ejecutar consulta
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Obtener el rol requerido
    If Not rs.EOF Then
        If IsNull(rs.Fields("RolRequerido").value) Then
            IWorkflowRepository_GetTransitionRequiredRole = ""
        Else
            IWorkflowRepository_GetTransitionRequiredRole = rs.Fields("RolRequerido").value
        End If
    Else
        IWorkflowRepository_GetTransitionRequiredRole = ""
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    IWorkflowRepository_GetTransitionRequiredRole = ""
    Debug.Print "Error en CWorkflowRepository.GetTransitionRequiredRole: " & Err.Number & " - " & Err.Description
End Function

'==============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'==============================================================================

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String) As Boolean
    IsValidTransition = IWorkflowRepository_IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetAvailableStates(ByVal tipoSolicitud As String) As Collection
    Set GetAvailableStates = IWorkflowRepository_GetAvailableStates(tipoSolicitud)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Collection
    Set GetNextStates = IWorkflowRepository_GetNextStates(estadoActual, tipoSolicitud, usuarioRol)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetInitialState(ByVal tipoSolicitud As String) As String
    GetInitialState = IWorkflowRepository_GetInitialState(tipoSolicitud)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function IsStateFinal(ByVal estadoCodigo As String, ByVal tipoSolicitud As String) As Boolean
    IsStateFinal = IWorkflowRepository_IsStateFinal(estadoCodigo, tipoSolicitud)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function RecordStateChange(ByVal SolicitudID As Long, ByVal estadoAnterior As String, ByVal estadoNuevo As String, ByVal usuario As String, ByVal comentarios As String) As Boolean
    RecordStateChange = IWorkflowRepository_RecordStateChange(SolicitudID, estadoAnterior, estadoNuevo, usuario, comentarios)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetStateHistory(ByVal SolicitudID As Long) As Collection
    Set GetStateHistory = IWorkflowRepository_GetStateHistory(SolicitudID)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function HasTransitionPermission(ByVal usuarioRol As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    HasTransitionPermission = IWorkflowRepository_HasTransitionPermission(usuarioRol, estadoOrigen, estadoDestino, tipoSolicitud)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function RequiresApproval(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As Boolean
    RequiresApproval = IWorkflowRepository_RequiresApproval(estadoOrigen, estadoDestino, tipoSolicitud)
End Function

' Método público de conveniencia que delega a la implementación de la interfaz
Public Function GetTransitionRequiredRole(ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String) As String
    GetTransitionRequiredRole = IWorkflowRepository_GetTransitionRequiredRole(estadoOrigen, estadoDestino, tipoSolicitud)
End Function


