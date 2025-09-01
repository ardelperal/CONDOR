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
' Version: 3.0 - Refactorización Total
' ============================================================================
' Implementa ISolicitudRepository con el nuevo contrato simplificado.
' Elimina el uso problemático de Collection y usa parámetros directos.
' ============================================================================

Implements ISolicitudRepository

' Dependencias inyectadas
Private m_configService As IConfig
Private m_errorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACIÓN Y CONFIGURACIÓN
' ============================================================================

' Inicializa el repositorio con sus dependencias
' @param configService: Servicio de configuración para obtener ruta del Backend
' @param errorHandler: Servicio de manejo de errores
Public Sub Initialize(ByVal configService As IConfig, ByVal errorHandler As IErrorHandlerService)
    If configService Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.Initialize", "configService no puede ser Nothing"
    End If
    
    If errorHandler Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.Initialize", "errorHandler no puede ser Nothing"
    End If
    
    Set m_configService = configService
    Set m_errorHandler = errorHandler
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE ISolicitudRepository
' ============================================================================

' Obtiene una solicitud por su ID desde el Backend
Private Function ISolicitudRepository_GetSolicitudById(ByVal idSolicitud As Long) As ESolicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.GetSolicitudById", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    Dim solicitud As ESolicitud
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Crear consulta parametrizada usando modQueries
    Set qdf = db.CreateQueryDef("", GET_SOLICITUD_BY_ID)
    
    ' Asignar parámetro directamente
    qdf.parameters(0) = idSolicitud
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = New ESolicitud
        With solicitud
            .idSolicitud = rs("idSolicitud")
            .idExpediente = rs("idExpediente")
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
    Else
        Set solicitud = Nothing
    End If
    
    rs.Close
    db.Close
    Set ISolicitudRepository_GetSolicitudById = solicitud
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    m_errorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.GetSolicitudById"
    Err.Raise Err.Number, "CSolicitudRepository.GetSolicitudById", Err.Description
End Function

' Guarda una solicitud en el Backend (INSERT o UPDATE)
Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "El repositorio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If solicitud Is Nothing Then
        Err.Raise 5, "CSolicitudRepository.SaveSolicitud", "solicitud no puede ser Nothing"
    End If
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim qdfIdentity As DAO.QueryDef
    Dim rs As DAO.recordset
    
    On Error GoTo ErrorHandler
    
    ' Conectar al Backend usando la ruta de configuración
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)
    
    ' Determinar si es INSERT o UPDATE basado en idSolicitud
    If solicitud.idSolicitud = 0 Then
        ' INSERT - Nueva solicitud
        Set qdf = db.CreateQueryDef("", INSERT_SOLICITUD)
        
        ' Asignar parámetros directamente desde el objeto solicitud
        qdf.parameters("idExpediente") = solicitud.idExpediente
        qdf.parameters("tipoSolicitud") = solicitud.tipoSolicitud
        qdf.parameters("subTipoSolicitud") = solicitud.subTipoSolicitud
        qdf.parameters("codigoSolicitud") = solicitud.codigoSolicitud
        qdf.parameters("idEstadoInterno") = solicitud.idEstadoInterno
        qdf.parameters("fechaCreacion") = solicitud.fechaCreacion
        qdf.parameters("usuarioCreacion") = solicitud.usuarioCreacion
        qdf.parameters("fechaModificacion") = solicitud.fechaModificacion
        qdf.parameters("usuarioModificacion") = solicitud.usuarioModificacion
        
        qdf.Execute dbFailOnError
        
        ' Obtener el ID generado
        Set qdfIdentity = db.CreateQueryDef("", GET_LAST_INSERT_ID)
        Set rs = qdfIdentity.OpenRecordset(dbOpenSnapshot)
        solicitud.idSolicitud = rs("LastID")
        rs.Close
    Else
        ' UPDATE - Solicitud existente
        Set qdf = db.CreateQueryDef("", UPDATE_SOLICITUD)
        
        ' Asignar parámetros directamente desde el objeto solicitud
        qdf.parameters("idExpediente") = solicitud.idExpediente
        qdf.parameters("tipoSolicitud") = solicitud.tipoSolicitud
        qdf.parameters("subTipoSolicitud") = solicitud.subTipoSolicitud
        qdf.parameters("codigoSolicitud") = solicitud.codigoSolicitud
        qdf.parameters("idEstadoInterno") = solicitud.idEstadoInterno
        qdf.parameters("fechaCreacion") = solicitud.fechaCreacion
        qdf.parameters("usuarioCreacion") = solicitud.usuarioCreacion
        qdf.parameters("fechaModificacion") = solicitud.fechaModificacion
        qdf.parameters("usuarioModificacion") = solicitud.usuarioModificacion
        qdf.parameters("idSolicitud") = solicitud.idSolicitud
        
        qdf.Execute dbFailOnError
    End If
    
    db.Close
    ISolicitudRepository_SaveSolicitud = solicitud.idSolicitud
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    m_errorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.SaveSolicitud"
    ISolicitudRepository_SaveSolicitud = 0
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

' Carga los datos específicos según el tipo de solicitud
Private Sub CargarDatosEspecificos(ByRef solicitud As ESolicitud, ByRef db As DAO.Database)
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.recordset
    
    On Error GoTo ErrorHandler
    
    Select Case UCase(solicitud.tipoSolicitud)
        Case "PC"
            Set qdf = db.CreateQueryDef("", GET_DATOS_PC_BY_SOLICITUD)
            qdf.parameters(0) = solicitud.idSolicitud
            Set rs = qdf.OpenRecordset(dbOpenSnapshot)
            
            If Not rs.EOF Then
                Dim datosPC As EDatosPc
                Set datosPC = New EDatosPc
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
                Dim datosCDCA As EDatosCdCa
                Set datosCDCA = New EDatosCdCa
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
                Dim datosCDCASUB As EDatosCdCaSub
                Set datosCDCASUB = New EDatosCdCaSub
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
    m_errorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.CargarDatosEspecificos"
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' Método público de conveniencia para GetSolicitudById
Public Function GetSolicitudById(ByVal idSolicitud As Long) As ESolicitud
    Set GetSolicitudById = ISolicitudRepository_GetSolicitudById(idSolicitud)
End Function

' Método público de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function


