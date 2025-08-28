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
Private Function ISolicitudRepository_GetSolicitudById(ByVal params As Collection) As E_Solicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.GetSolicitudById", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim solicitud As E_Solicitud
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada usando modQueries
    Set qdf = db.CreateQueryDef("", GET_SOLICITUD_BY_ID)
    
    ' Asignar parámetros desde la Collection
    Dim i As Integer
    Dim param As QueryParameter
    i = 0
    For Each param In params
        qdf.parameters(i) = param.ParameterValue
        i = i + 1
    Next param
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = New E_Solicitud
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
Private Function ISolicitudRepository_SaveSolicitud(ByVal params As Collection) As Long
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If params Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "params no puede ser Nothing"
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim qdfIdentity As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim resultId As Long
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Determinar si es INSERT o UPDATE basado en el primer parámetro (idSolicitud)
    Dim firstParam As QueryParameter
    Set firstParam = params(1)
    
    If firstParam.ParameterValue = 0 Then
        ' INSERT - Nueva solicitud usando consulta centralizada
        Set qdf = db.CreateQueryDef("", INSERT_SOLICITUD)
        
        ' Asignar parámetros desde la Collection (saltando el idSolicitud para INSERT)
        Dim i As Integer
        Dim param As QueryParameter
        i = 0
        For Each param In params
            If i > 0 Then ' Saltar el primer parámetro (idSolicitud) para INSERT
                qdf.parameters(i - 1) = param.ParameterValue
            End If
            i = i + 1
        Next param
        
        qdf.Execute dbFailOnError
        
        ' Obtener el ID generado usando consulta centralizada
        Set qdfIdentity = db.CreateQueryDef("", GET_LAST_INSERT_ID)
        Set rs = qdfIdentity.OpenRecordset(dbOpenSnapshot)
        resultId = rs("LastID")
        rs.Close
        
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    Else
        ' UPDATE - Solicitud existente usando consulta centralizada
        Set qdf = db.CreateQueryDef("", UPDATE_SOLICITUD)
        
        ' Asignar parámetros desde la Collection
        i = 0
        For Each param In params
            qdf.parameters(i) = param.ParameterValue
            i = i + 1
        Next param
        
        qdf.Execute dbFailOnError
        
        resultId = firstParam.ParameterValue
        
        ' Logging de auditoría removido - responsabilidad de la capa de servicio
    End If
    
    db.Close
    ISolicitudRepository_SaveSolicitud = resultId
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    ' Logging de errores removido - responsabilidad de la capa de servicio
    Err.Raise Err.Number, "CSolicitudRepository.SaveSolicitud", Err.Description
End Function

' Ejecuta una consulta SQL con parámetros
Private Function ISolicitudRepository_ExecuteQuery(ByVal queryName As String, ByVal params As Collection) As DAO.recordset
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.ExecuteQuery", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Obtener la consulta SQL desde modQueries usando el nombre
    sql = CallByName(modQueries, queryName, VbGet)
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada
    Set qdf = db.CreateQueryDef("", sql)
    
    ' Aplicar parámetros desde la Collection
    Dim i As Integer
    Dim param As QueryParameter
    i = 0
    For Each param In params
        qdf.parameters(i) = param.ParameterValue
        i = i + 1
    Next param
    
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
Private Sub CargarDatosEspecificos(ByRef solicitud As E_Solicitud, ByRef db As DAO.Database)
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    On Error GoTo ErrorHandler
    
    Select Case UCase(solicitud.tipoSolicitud)
        Case "PC"
            Set qdf = db.CreateQueryDef("", GET_DATOS_PC_BY_SOLICITUD)
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosPC As E_Datos_PC
Set datosPC = New E_Datos_PC
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
            Set qdf = db.CreateQueryDef("", GET_DATOS_CD_CA_BY_SOLICITUD)
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosCDCA As E_Datos_CD_CA
Set datosCDCA = New E_Datos_CD_CA
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
            Set qdf = db.CreateQueryDef("", GET_DATOS_CD_CA_SUB_BY_SOLICITUD)
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosCDCASUB As E_Datos_CD_CA_SUB
Set datosCDCASUB = New E_Datos_CD_CA_SUB
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
Public Function GetSolicitudById(ByVal idSolicitud As Long) As E_Solicitud
    Dim params As Collection
    Set params = New Collection
    
    Dim param As QueryParameter
    Set param = New QueryParameter
    param.ParameterName = "idSolicitud"
    param.ParameterValue = idSolicitud
    param.DataType = dbLong
    params.Add param
    
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(params)
End Function

' Método público de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As E_Solicitud) As Long
    Dim params As Collection
    Set params = New Collection
    
    ' Construir parámetros desde el objeto solicitud
    Dim param As QueryParameter
    
    ' idSolicitud (0 para INSERT, valor existente para UPDATE)
    Set param = New QueryParameter
    param.ParameterName = "idSolicitud"
    param.ParameterValue = solicitud.idSolicitud
    param.DataType = dbLong
    params.Add param
    
    ' idExpediente
    Set param = New QueryParameter
    param.ParameterName = "idExpediente"
    param.ParameterValue = solicitud.idExpediente
    param.DataType = dbText
    params.Add param
    
    ' tipoSolicitud
    Set param = New QueryParameter
    param.ParameterName = "tipoSolicitud"
    param.ParameterValue = solicitud.tipoSolicitud
    param.DataType = dbText
    params.Add param
    
    ' subTipoSolicitud
    Set param = New QueryParameter
    param.ParameterName = "subTipoSolicitud"
    param.ParameterValue = solicitud.subTipoSolicitud
    param.DataType = dbText
    params.Add param
    
    ' codigoSolicitud
    Set param = New QueryParameter
    param.ParameterName = "codigoSolicitud"
    param.ParameterValue = solicitud.codigoSolicitud
    param.DataType = dbText
    params.Add param
    
    ' idEstadoInterno
    Set param = New QueryParameter
    param.ParameterName = "idEstadoInterno"
    param.ParameterValue = solicitud.idEstadoInterno
    param.DataType = dbInteger
    params.Add param
    
    ' fechaCreacion
    Set param = New QueryParameter
    param.ParameterName = "fechaCreacion"
    param.ParameterValue = solicitud.fechaCreacion
    param.DataType = dbDate
    params.Add param
    
    ' usuarioCreacion
    Set param = New QueryParameter
    param.ParameterName = "usuarioCreacion"
    param.ParameterValue = solicitud.usuarioCreacion
    param.DataType = dbText
    params.Add param
    
    ' fechaModificacion
    Set param = New QueryParameter
    param.ParameterName = "fechaModificacion"
    param.ParameterValue = solicitud.fechaModificacion
    param.DataType = dbDate
    params.Add param
    
    ' usuarioModificacion
    Set param = New QueryParameter
    param.ParameterName = "usuarioModificacion"
    param.ParameterValue = solicitud.usuarioModificacion
    param.DataType = dbText
    params.Add param
    
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud, params)
End Function

' Método público de conveniencia para ExecuteQuery
Public Function ExecuteQuery(ByVal queryName As String, Optional ByVal params As Collection = Nothing) As DAO.recordset
    If params Is Nothing Then
        Set params = New Collection
    End If
    Set ExecuteQuery = ISolicitudRepository_ExecuteQuery(queryName, params)
End Function


