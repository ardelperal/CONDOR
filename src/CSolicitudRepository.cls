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
Private m_ErrorHandler As IErrorHandlerService
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
    Set m_ErrorHandler = errorHandler
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE ISolicitudRepository
' ============================================================================



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
    Dim rs As DAO.Recordset
    
    On Error GoTo errorHandler
    
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
        qdf.Parameters("pIdExpediente") = solicitud.idExpediente
        qdf.Parameters("pTipoSolicitud") = solicitud.tipoSolicitud
        qdf.Parameters("pSubTipoSolicitud") = solicitud.subTipoSolicitud
        qdf.Parameters("pCodigoSolicitud") = solicitud.codigoSolicitud
        qdf.Parameters("pIdEstadoInterno") = solicitud.idEstadoInterno
        qdf.Parameters("pUsuarioCreacion") = solicitud.usuarioCreacion
        
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
        qdf.Parameters("pIdExpediente") = solicitud.idExpediente
        qdf.Parameters("pTipoSolicitud") = solicitud.tipoSolicitud
        qdf.Parameters("pSubTipoSolicitud") = solicitud.subTipoSolicitud
        qdf.Parameters("pCodigoSolicitud") = solicitud.codigoSolicitud
        qdf.Parameters("pIdEstadoInterno") = solicitud.idEstadoInterno
        qdf.Parameters("pUsuarioModificacion") = solicitud.usuarioModificacion
        qdf.Parameters("pIdSolicitud") = solicitud.idSolicitud
        
        qdf.Execute dbFailOnError
    End If
    
    db.Close
    ISolicitudRepository_SaveSolicitud = solicitud.idSolicitud
    Exit Function
    
errorHandler:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.SaveSolicitud"
    ISolicitudRepository_SaveSolicitud = 0
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

' Carga los datos específicos según el tipo de solicitud
Private Sub CargarDatosEspecificos(ByRef solicitud As ESolicitud, ByRef db As DAO.Database)
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    On Error GoTo errorHandler
    
    Select Case UCase(solicitud.tipoSolicitud)
        Case "PC"
            Set qdf = db.CreateQueryDef("", GET_DATOS_PC_BY_SOLICITUD)
            qdf.Parameters(0) = solicitud.idSolicitud
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
            Set qdf = db.CreateQueryDef("", GET_DATOS_CDCA_BY_SOLICITUD)
            qdf.Parameters(0) = solicitud.idSolicitud
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
            Set qdf = db.CreateQueryDef("", GET_DATOS_CDCASUB_BY_SOLICITUD)
            qdf.Parameters(0) = solicitud.idSolicitud
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
    
errorHandler:
    If Not rs Is Nothing Then rs.Close
    ' Log del error pero no propagar para no interrumpir la carga principal
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.CargarDatosEspecificos"
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================
' Estos métodos exponen públicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta



' Método público de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function

Private Function ISolicitudRepository_ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Dim solicitud As ESolicitud ' Declaración al principio
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    On Error GoTo errorHandler
    
    ' Conexión a la BD
    Set db = DBEngine.OpenDatabase(m_configService.GetDataPath(), False, False, ";PWD=" & m_configService.GetDatabasePassword())
    
    ' Preparar la consulta parametrizada
    Set qdf = db.CreateQueryDef("", GET_SOLICITUD_BY_ID)
    qdf.Parameters("pIdSolicitud") = SolicitudID
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' CORRECCIÓN CRÍTICA: Instanciar el objeto ANTES de usarlo
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
        
        ' Cargar datos específicos después de que el objeto principal está poblado
        Call CargarDatosEspecificos(solicitud, db)
    Else
        Set solicitud = Nothing ' Explícitamente Nothing si no se encuentra
    End If
    
    ' Asignar el valor de retorno ANTES de cerrar los recursos
    Set ISolicitudRepository_ObtenerSolicitudPorId = solicitud
    
Cleanup:
    ' Limpieza de recursos garantizada
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function
    
errorHandler:
    ' Loguear el error y devolver Nothing
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudRepository.ObtenerSolicitudPorId"
    Set ISolicitudRepository_ObtenerSolicitudPorId = Nothing
    Resume Cleanup
End Function

' Método público de conveniencia para ObtenerSolicitudPorId
Public Function ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Set ObtenerSolicitudPorId = ISolicitudRepository_ObtenerSolicitudPorId(SolicitudID)
End Function




