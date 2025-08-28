Attribute VB_Name = "modQueries"
Option Compare Database
Option Explicit

'''
' Módulo Central de Consultas SQL - CONDOR
' Contiene todas las sentencias SQL parametrizadas del sistema
' Principio: Centralización de consultas para mantenibilidad y seguridad
'''

' ============================================================================
' CONSULTAS DE EXPEDIENTES
' ============================================================================

Public Const GET_EXPEDIENTE_BY_ID As String = _
    "SELECT E.IDExpediente, E.Nemotecnico, E.FechaCreacion, E.UsuarioCreacion, " & _
    "E.FechaModificacion, E.UsuarioModificacion, E.Activo " & _
    "FROM TbExpedientes AS E " & _
    "WHERE E.IDExpediente = [pIdExpediente];"

Public Const GET_EXPEDIENTE_BY_NEMOTECNICO As String = _
    "SELECT E.IDExpediente, E.Nemotecnico, E.FechaCreacion, E.UsuarioCreacion, " & _
    "E.FechaModificacion, E.UsuarioModificacion, E.Activo " & _
    "FROM TbExpedientes AS E " & _
    "WHERE E.Nemotecnico = [pNemotecnico];"

Public Const GET_EXPEDIENTES_ACTIVOS_SELECTOR As String = _
    "SELECT E.IDExpediente, E.Nemotecnico " & _
    "FROM TbExpedientes AS E " & _
    "WHERE E.Activo = True " & _
    "ORDER BY E.Nemotecnico;"

' ============================================================================
' CONSULTAS DE AUTENTICACIÓN
' ============================================================================

Public Const GET_AUTH_DATA_BY_EMAIL As String = _
    "SELECT ua.EsAdministrador, uap.EsUsuarioCalidad, ua.CorreoElectronico " & _
    "FROM TbUsuariosAplicaciones ua " & _
    "LEFT JOIN TbUsuariosAplicacionesPerfiles uap ON ua.ID = uap.UsuarioID " & _
    "WHERE ua.CorreoElectronico = [pEmail];"

' ============================================================================
' CONSULTAS DE WORKFLOW
' ============================================================================

Public Const IS_VALID_TRANSITION As String = _
    "SELECT COUNT(*) AS TransitionCount " & _
    "FROM TbTransiciones T " & _
    "INNER JOIN TbEstados EO ON T.EstadoOrigenID = EO.ID " & _
    "INNER JOIN TbEstados ED ON T.EstadoDestinoID = ED.ID " & _
    "WHERE EO.CodigoEstado = [pEstadoOrigen] " & _
    "AND ED.CodigoEstado = [pEstadoDestino] " & _
    "AND T.TipoSolicitud = [pTipoSolicitud] " & _
    "AND T.Activo = True;"

Public Const GET_AVAILABLE_STATES As String = _
    "SELECT E.CodigoEstado, E.NombreEstado " & _
    "FROM TbEstados E " & _
    "INNER JOIN TbTiposSolicitud TS ON E.TipoSolicitud = TS.Codigo " & _
    "WHERE TS.Codigo = [pTipoSolicitud] " & _
    "AND E.Activo = True " & _
    "ORDER BY E.NombreEstado;"

Public Const GET_NEXT_STATES As String = _
    "SELECT ED.CodigoEstado, ED.NombreEstado " & _
    "FROM TbTransiciones T " & _
    "INNER JOIN TbEstados EO ON T.EstadoOrigenID = EO.ID " & _
    "INNER JOIN TbEstados ED ON T.EstadoDestinoID = ED.ID " & _
    "WHERE EO.CodigoEstado = [pEstadoActual] " & _
    "AND T.TipoSolicitud = [pTipoSolicitud] " & _
    "AND T.Activo = True " & _
    "ORDER BY ED.NombreEstado;"

Public Const GET_INITIAL_STATE As String = _
    "SELECT E.CodigoEstado " & _
    "FROM TbEstados E " & _
    "INNER JOIN TbTiposSolicitud TS ON E.ID = TS.EstadoInicialID " & _
    "WHERE TS.Codigo = [pTipoSolicitud];"

Public Const IS_STATE_FINAL As String = _
    "SELECT COUNT(*) AS TransitionCount " & _
    "FROM TbTransiciones T " & _
    "INNER JOIN TbEstados E ON T.EstadoOrigenID = E.ID " & _
    "WHERE E.CodigoEstado = [pEstadoCodigo] " & _
    "AND T.TipoSolicitud = [pTipoSolicitud] " & _
    "AND T.Activo = True;"

Public Const RECORD_STATE_CHANGE As String = _
    "INSERT INTO TbHistorialEstados (SolicitudID, EstadoAnterior, EstadoNuevo, Usuario, FechaCambio, Comentarios) " & _
    "VALUES ([pSolicitudID], [pEstadoAnterior], [pEstadoNuevo], [pUsuario], Now(), [pComentarios]);"

Public Const GET_STATE_HISTORY As String = _
    "SELECT EstadoAnterior, EstadoNuevo, Usuario, FechaCambio, Comentarios " & _
    "FROM TbHistorialEstados " & _
    "WHERE SolicitudID = [pSolicitudID] " & _
    "ORDER BY FechaCambio DESC;"

' ============================================================================
' CONSULTAS DE SOLICITUDES
' ============================================================================

Public Const GET_SOLICITUD_BY_ID As String = _
    "SELECT * FROM T_Solicitudes WHERE idSolicitud = [pIdSolicitud];"

Public Const INSERT_SOLICITUD As String = _
    "INSERT INTO T_Solicitudes (idExpediente, tipoSolicitud, subTipoSolicitud, " & _
    "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion) " & _
    "VALUES ([pIdExpediente], [pTipoSolicitud], [pSubTipoSolicitud], " & _
    "[pCodigoSolicitud], [pIdEstadoInterno], [pFechaCreacion], [pUsuarioCreacion]);"

Public Const UPDATE_SOLICITUD As String = _
    "UPDATE T_Solicitudes SET idExpediente=[pIdExpediente], tipoSolicitud=[pTipoSolicitud], " & _
    "subTipoSolicitud=[pSubTipoSolicitud], codigoSolicitud=[pCodigoSolicitud], " & _
    "idEstadoInterno=[pIdEstadoInterno], fechaModificacion=[pFechaModificacion], " & _
    "usuarioModificacion=[pUsuarioModificacion] " & _
    "WHERE idSolicitud=[pIdSolicitud];"

Public Const GET_DATOS_PC_BY_SOLICITUD As String = _
    "SELECT * FROM TbDatos_PC WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_DATOS_CD_CA_BY_SOLICITUD As String = _
    "SELECT * FROM TbDatos_CD_CA WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_DATOS_CD_CA_SUB_BY_SOLICITUD As String = _
    "SELECT * FROM TbDatos_CD_CA_SUB WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_LAST_INSERT_ID As String = _
    "SELECT @@IDENTITY AS LastID;"

Public Const GET_IDENTITY As String = _
    "SELECT @@IDENTITY;"

' ============================================================================
' CONSULTAS PARA CONFIGURACIÓN (TbConfiguracion)
' ============================================================================

' Consulta para obtener toda la configuración del sistema
Public Const GET_ALL_CONFIGURATION As String = "SELECT Clave, Valor FROM TbConfiguracion;"

' ============================================================================
' CONSULTAS DE MAPEO
' ============================================================================

Public Const GET_MAPEO_POR_TIPO As String = _
    "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado " & _
    "FROM tbMapeoCampos " & _
    "WHERE nombrePlantilla = [pTipoSolicitud];"

' ============================================================================
' CONSULTAS DE CONFIGURACIÓN
' ============================================================================

Public Const GET_CONFIG_VALUE As String = _
    "SELECT ConfigValue FROM TbConfiguracion WHERE ConfigKey = [pConfigKey];"

Public Const GET_ALL_CONFIG As String = _
    "SELECT ConfigKey, ConfigValue FROM TbConfiguracion ORDER BY ConfigKey;"