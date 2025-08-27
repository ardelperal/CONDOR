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
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACIÓN Y CONFIGURACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta del Backend
Public Sub Initialize(ByVal configService As IConfig)
    If configService Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    Set m_configService = configService
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE ISolicitudRepository
' ============================================================================

' Obtiene una solicitud por su ID desde el Backend
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.GetSolicitudById", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim solicitud As T_Solicitud
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada para evitar inyección SQL
    Set qdf = db.CreateQueryDef("", "SELECT * FROM T_Solicitudes WHERE idSolicitud = ?")
    qdf.parameters(0) = idSolicitud
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = New T_Solicitud
        With solicitud
            .idSolicitud = rs("idSolicitud")
            .idExpediente = rs("idExpediente") & ""
            .tipoSolicitud = rs("tipoSolicitud") & ""
            .subTipoSolicitud = rs("subTipoSolicitud") & ""
            .codigoSolicitud = rs("codigoSolicitud") & ""
            .idEstadoInterno = rs("idEstadoInterno")
            .fechaCreacion = rs("fechaCreacion")
            .usuarioCreacion = rs("usuarioCreacion") & ""
            .fechaModificacion = rs("fechaModificacion")
            .usuarioModificacion = rs("usuarioModificacion") & ""
        End With
        
        ' Cargar datos específicos según el tipo de solicitud
        Call CargarDatosEspecificos(solicitud, db)
        
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    Else
        Set solicitud = Nothing
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    End If
    
    rs.Close
    db.Close
    Set ISolicitudRepository_GetSolicitudById = solicitud
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    ' Logging de errores removido - responsabilidad de la capa de servicio
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
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    If solicitud.idSolicitud = 0 Then
        ' INSERT - Nueva solicitud
        Set qdf = db.CreateQueryDef("", _
            "INSERT INTO T_Solicitudes (idExpediente, tipoSolicitud, subTipoSolicitud, " & _
            "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?)")
        
        qdf.parameters(0) = solicitud.idExpediente
        qdf.parameters(1) = solicitud.tipoSolicitud
        qdf.parameters(2) = solicitud.subTipoSolicitud
        qdf.parameters(3) = solicitud.codigoSolicitud
        qdf.parameters(4) = solicitud.idEstadoInterno
        qdf.parameters(5) = solicitud.fechaCreacion
        qdf.parameters(6) = solicitud.usuarioCreacion
        
        qdf.Execute dbFailOnError
        
        ' Obtener el ID generado
        resultId = db.OpenRecordset("SELECT @@IDENTITY")(0)
        solicitud.idSolicitud = resultId
        
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    Else
        ' UPDATE - Solicitud existente
        Set qdf = db.CreateQueryDef("", _
            "UPDATE T_Solicitudes SET idExpediente=?, tipoSolicitud=?, subTipoSolicitud=?, " & _
            "codigoSolicitud=?, idEstadoInterno=?, fechaModificacion=?, usuarioModificacion=? " & _
            "WHERE idSolicitud=?")
        
        qdf.parameters(0) = solicitud.idExpediente
        qdf.parameters(1) = solicitud.tipoSolicitud
        qdf.parameters(2) = solicitud.subTipoSolicitud
        qdf.parameters(3) = solicitud.codigoSolicitud
        qdf.parameters(4) = solicitud.idEstadoInterno
        qdf.parameters(5) = Now()
        qdf.parameters(6) = solicitud.usuarioModificacion
        qdf.parameters(7) = solicitud.idSolicitud
        
        qdf.Execute dbFailOnError
        
        resultId = solicitud.idSolicitud
        
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    End If
    
    db.Close
    ISolicitudRepository_SaveSolicitud = resultId
    Exit Function
    
ErrorHandler:
    If Not db Is Nothing Then db.Close
    ' Logging de errores removido - responsabilidad de la capa de servicio
    Err.Raise Err.Number, "CSolicitudRepository.SaveSolicitud", Err.Description
End Function

' Ejecuta una consulta SQL con parámetros opcionales
Private Function ISolicitudRepository_ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.recordset
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.ExecuteQuery", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada
    Set qdf = db.CreateQueryDef("", sql)
    
    ' Aplicar parámetros si se proporcionaron
    If Not params Is Nothing Then
        Dim i As Integer
        Dim param As QueryParameter
        i = 0
        For Each param In params
            qdf.parameters(i) = param.ParameterValue
            i = i + 1
        Next param
    End If
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Logging de auditoría removido - responsabilidad de la capa de servicio
    
    Set ISolicitudRepository_ExecuteQuery = rs
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    ' Logging de errores removido - responsabilidad de la capa de servicio
    Err.Raise Err.Number, "CSolicitudRepository.ExecuteQuery", Err.Description
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

' Carga los datos específicos según el tipo de solicitud
Private Sub CargarDatosEspecificos(ByRef solicitud As T_Solicitud, ByRef db As DAO.Database)
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    On Error GoTo ErrorHandler
    
    Select Case UCase(solicitud.tipoSolicitud)
        Case "PC"
            Set qdf = db.CreateQueryDef("", "SELECT * FROM TbDatos_PC WHERE idSolicitud = ?")
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosPC As T_Datos_PC
                Set datosPC = New T_Datos_PC
                ' Cargar propiedades de datosPC desde el recordset
                With datosPC
                    .idDatosPC = rs("idDatosPC")
                    .idSolicitud = rs("idSolicitud")
                    .refContratoInspeccionOficial = rs("refContratoInspeccionOficial") & ""
                    .refSuministrador = rs("refSuministrador") & ""
                    .suministradorNombreDir = rs("suministradorNombreDir") & ""
                    .objetoContrato = rs("objetoContrato") & ""
                    .descripcionMaterialAfectado = rs("descripcionMaterialAfectado") & ""
                    .numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
                    .descripcionPropuestaCambio = rs("descripcionPropuestaCambio") & ""
                End With
                Set solicitud.datosPC = datosPC
            End If
            rs.Close
            
        Case "CD_CA"
            Set qdf = db.CreateQueryDef("", "SELECT * FROM TbDatos_CD_CA WHERE idSolicitud = ?")
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosCDCA As T_Datos_CD_CA
                Set datosCDCA = New T_Datos_CD_CA
                ' Cargar propiedades de datosCDCA desde el recordset
                With datosCDCA
                    .idDatosCDCA = rs("idDatosCDCA")
                    .idSolicitud = rs("idSolicitud")
                    .refSuministrador = rs("refSuministrador") & ""
                    .numContrato = rs("numContrato") & ""
                    .identificacionMaterial = rs("identificacionMaterial") & ""
                    .numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
                    .cantidadPeriodo = rs("cantidadPeriodo") & ""
                End With
                Set solicitud.datosCDCA = datosCDCA
            End If
            rs.Close
            
        Case "CD_CA_SUB"
            Set qdf = db.CreateQueryDef("", "SELECT * FROM TbDatos_CD_CA_SUB WHERE idSolicitud = ?")
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosCDCASUB As T_Datos_CD_CA_SUB
                Set datosCDCASUB = New T_Datos_CD_CA_SUB
                ' Cargar propiedades de datosCDCASUB desde el recordset
                With datosCDCASUB
                    .idDatosCDCASUB = rs("idDatosCDCASUB")
                    .idSolicitud = rs("idSolicitud")
                    .refSuministrador = rs("refSuministrador") & ""
                    .refSubSuministrador = rs("refSubSuministrador") & ""
                    .suministradorPrincipalNombreDir = rs("suministradorPrincipalNombreDir") & ""
                    .subSuministradorNombreDir = rs("subSuministradorNombreDir") & ""
                    .identificacionMaterial = rs("identificacionMaterial") & ""
                    .numPlanoEspecificacion = rs("numPlanoEspecificacion") & ""
                    .cantidadPeriodo = rs("cantidadPeriodo")
                    .numSerieLote = rs("numSerieLote") & ""
                End With
                Set solicitud.datosCDCASUB = datosCDCASUB
            End If
            rs.Close
    End Select
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    ' Log del error pero no propagar para no interrumpir la carga principal
    ' Logging de errores removido - responsabilidad de la capa de servicio
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para GetSolicitudById
Public Function GetSolicitudById(ByVal idSolicitud As Long) As T_Solicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

' Método público de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As T_Solicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function

' Método público de conveniencia para ExecuteQuery
Public Function ExecuteQuery(ByVal sql As String, Optional ByVal params As Collection = Nothing) As DAO.recordset
    Set ExecuteQuery = ISolicitudRepository_ExecuteQuery(sql, params)
End Function


