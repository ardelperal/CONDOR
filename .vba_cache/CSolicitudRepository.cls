VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' CLASE: CSolicitudRepository
' Arquitectura: Implementación del Repositorio de Solicitudes
' Version: 2.0 - Reconstrucción Total
' ============================================================================
' Implementa ISolicitudRepository con conexión explícita al Backend.
' Aplica principios de inyección de dependencias y consultas parametrizadas.
' ============================================================================

Implements ISolicitudRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_operationLogger As IOperationLogger
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACIÓN Y CONFIGURACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta del Backend
' @param operationLogger: Logger para registrar operaciones
Public Sub Initialize(ByVal configService As IConfig, ByVal operationLogger As IOperationLogger)
    If configService Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    If operationLogger Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.Initialize", "operationLogger no puede ser Nothing"
    End If
    
    Set m_configService = configService
    Set m_operationLogger = operationLogger
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE ISolicitudRepository
' ============================================================================

' Obtiene una solicitud por su ID desde el Backend
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.GetSolicitudById", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim solicitud As T_Solicitud
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetValue("DATAPATH")
    
    Dim backendPassword As String
    backendPassword = m_configService.GetValue("DATAPASSWORD")
    
    Set db = DBEngine.OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada para evitar inyección SQL
    Set qdf = db.CreateQueryDef("", "SELECT * FROM T_Solicitudes WHERE idSolicitud = ?")
    qdf.Parameters(0) = idSolicitud
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = New T_Solicitud
        With solicitud
            .idSolicitud = rs("idSolicitud")
            .idExpediente = rs("idExpediente") & ""
            .tipoSolicitud = rs("tipoSolicitud") & ""
            .subTipoSolicitud = rs("subTipoSolicitud") & ""
            .codigoSolicitud = rs("codigoSolicitud") & ""
            .estadoInterno = rs("estadoInterno") & ""
            .fechaCreacion = rs("fechaCreacion")
            .usuarioCreacion = rs("usuarioCreacion") & ""
            .fechaModificacion = rs("fechaModificacion")
            .usuarioModificacion = rs("usuarioModificacion") & ""
        End With
        
        m_operationLogger.LogOperation "CSolicitudRepository.GetSolicitudById", "Solicitud obtenida: ID=" & idSolicitud
    Else
        Set solicitud = Nothing
        m_operationLogger.LogOperation "CSolicitudRepository.GetSolicitudById", "Solicitud no encontrada: ID=" & idSolicitud
    End If
    
    rs.Close
    db.Close
    Set ISolicitudRepository_GetSolicitudById = solicitud
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    m_operationLogger.LogOperation "CSolicitudRepository.GetSolicitudById", "Error: " & Err.Description
    Err.Raise Err.Number, "CSolicitudRepository.GetSolicitudById", Err.Description
End Function

' Guarda una solicitud en el Backend (INSERT o UPDATE)
Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If solicitud Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "solicitud no puede ser Nothing"
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim resultId As Long
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetValue("DATAPATH")
    
    Dim backendPassword As String
    backendPassword = m_configService.GetValue("DATAPASSWORD")
    
    Set db = DBEngine.OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
    
    If solicitud.idSolicitud = 0 Then
        ' INSERT - Nueva solicitud
        Set qdf = db.CreateQueryDef("", _
            "INSERT INTO T_Solicitudes (idExpediente, tipoSolicitud, subTipoSolicitud, " & _
            "codigoSolicitud, estadoInterno, fechaCreacion, usuarioCreacion) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?)")
        
        qdf.Parameters(0) = solicitud.idExpediente
        qdf.Parameters(1) = solicitud.tipoSolicitud
        qdf.Parameters(2) = solicitud.subTipoSolicitud
        qdf.Parameters(3) = solicitud.codigoSolicitud
        qdf.Parameters(4) = solicitud.estadoInterno
        qdf.Parameters(5) = solicitud.fechaCreacion
        qdf.Parameters(6) = solicitud.usuarioCreacion
        
        qdf.Execute dbFailOnError
        
        ' Obtener el ID generado
        resultId = db.OpenRecordset("SELECT @@IDENTITY")(0)
        solicitud.idSolicitud = resultId
        
        m_operationLogger.LogOperation "CSolicitudRepository.SaveSolicitud", "Nueva solicitud creada: ID=" & resultId
    Else
        ' UPDATE - Solicitud existente
        Set qdf = db.CreateQueryDef("", _
            "UPDATE T_Solicitudes SET idExpediente=?, tipoSolicitud=?, subTipoSolicitud=?, " & _
            "codigoSolicitud=?, estadoInterno=?, fechaModificacion=?, usuarioModificacion=? " & _
            "WHERE idSolicitud=?")
        
        qdf.Parameters(0) = solicitud.idExpediente
        qdf.Parameters(1) = solicitud.tipoSolicitud
        qdf.Parameters(2) = solicitud.subTipoSolicitud
        qdf.Parameters(3) = solicitud.codigoSolicitud
        qdf.Parameters(4) = solicitud.estadoInterno
        qdf.Parameters(5) = Now()
        qdf.Parameters(6) = solicitud.usuarioModificacion
        qdf.Parameters(7) = solicitud.idSolicitud
        
        qdf.Execute dbFailOnError
        
        resultId = solicitud.idSolicitud
        
        m_operationLogger.LogOperation "CSolicitudRepository.SaveSolicitud", "Solicitud actualizada: ID=" & resultId
    End If
    
    db.Close
    ISolicitudRepository_SaveSolicitud = resultId
    Exit Function
    
ErrorHandler:
    If Not db Is Nothing Then db.Close
    m_operationLogger.LogOperation "CSolicitudRepository.SaveSolicitud", "Error: " & Err.Description
    Err.Raise Err.Number, "CSolicitudRepository.SaveSolicitud", Err.Description
End Function