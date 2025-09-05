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

Implements ISolicitudRepository

Private Const MODULE_NAME As String = "CSolicitudRepository"

Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_ErrorHandler = ErrorHandler
End Sub

' --- IMPLEMENTACIÓN DE LA INTERFAZ ---

Private Function ISolicitudRepository_ObtenerSolicitudPorId(ByVal ID As Long) As ESolicitud
    Dim db As DAO.Database, qdf As DAO.QueryDef, rs As DAO.Recordset
    Dim solicitud As ESolicitud
    On Error GoTo ErrorHandler
    
    Set db = GetDatabaseConnection()
    If db Is Nothing Then
        ' Si la conexión falla, la función de log ya fue llamada. Salir de forma segura.
        Set ISolicitudRepository_ObtenerSolicitudPorId = Nothing
        Exit Function
    End If
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM tbSolicitudes WHERE idSolicitud = [pId]")
    qdf.Parameters("pId").value = ID
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set solicitud = MapRecordsetToSolicitud(rs)
    End If
    
    Set ISolicitudRepository_ObtenerSolicitudPorId = solicitud

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then qdf.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing: Set qdf = Nothing: Set db = Nothing
    Exit Function
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".ObtenerSolicitudPorId: " & Err.Description
    Resume Cleanup
End Function

Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim newId As Long
    newId = 0
    On Error GoTo ErrorHandler
    
    Set db = GetDatabaseConnection()
    
    If solicitud.idSolicitud = 0 Then ' Es un nuevo registro (INSERT)
        Set rs = db.OpenRecordset("tbSolicitudes", dbOpenDynaset, dbAppendOnly)
        rs.AddNew
    Else ' Es una actualización (UPDATE)
        Set rs = db.OpenRecordset("SELECT * FROM tbSolicitudes WHERE idSolicitud=" & solicitud.idSolicitud, dbOpenDynaset)
        If rs.EOF Then
            Err.Raise vbObjectError + 1002, MODULE_NAME, "No se encontró la solicitud con ID " & solicitud.idSolicitud & " para actualizar."
        End If
        rs.Edit
    End If
    
    ' Mapeo de objeto a recordset
    rs!idExpediente = solicitud.idExpediente
    rs!tipoSolicitud = solicitud.tipoSolicitud
    rs!subTipoSolicitud = solicitud.subTipoSolicitud
    rs!codigoSolicitud = solicitud.codigoSolicitud
    rs!idEstadoInterno = solicitud.idEstadoInterno
    rs!fechaCreacion = solicitud.fechaCreacion
    rs!usuarioCreacion = solicitud.usuarioCreacion
    rs!fechaPaseTecnico = solicitud.fechaPaseTecnico
    rs!fechaCompletadoTecnico = solicitud.fechaCompletadoTecnico
    rs!fechaModificacion = solicitud.fechaModificacion
    rs!usuarioModificacion = solicitud.usuarioModificacion

    
    rs.Update
    
    ' Obtener el ID para devolverlo
    If solicitud.idSolicitud = 0 Then
        rs.Bookmark = rs.LastModified
        newId = rs!idSolicitud
        solicitud.idSolicitud = newId ' Actualizar el objeto original con el nuevo ID
    Else
        newId = solicitud.idSolicitud
    End If
    
    ISolicitudRepository_SaveSolicitud = newId

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function
ErrorHandler:
    ISolicitudRepository_SaveSolicitud = 0 ' Devolver 0 en caso de error
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".SaveSolicitud: " & Err.Description
    Resume Cleanup
End Function

' --- FUNCIONES AUXILIARES ---

Private Function GetDatabaseConnection() As DAO.Database
    On Error GoTo ErrorHandler
    Dim dbPath As String: dbPath = m_Config.GetDataPath()
    Dim dbPassword As String: dbPassword = m_Config.GetDatabasePassword()
    Set GetDatabaseConnection = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & dbPassword)
    Exit Function
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".GetDatabaseConnection: " & Err.Description
End Function

Private Function MapRecordsetToSolicitud(ByVal rs As DAO.Recordset) As ESolicitud
    Dim s As New ESolicitud
    s.idSolicitud = Nz(rs!idSolicitud, 0)
    s.idExpediente = Nz(rs!idExpediente, 0)
    s.tipoSolicitud = Nz(rs!tipoSolicitud, "")
    s.subTipoSolicitud = Nz(rs!subTipoSolicitud, "")
    s.codigoSolicitud = Nz(rs!codigoSolicitud, "")
    s.idEstadoInterno = Nz(rs!idEstadoInterno, 0) ' Usando el campo Long normalizado
    s.fechaCreacion = rs!fechaCreacion
    s.usuarioCreacion = Nz(rs!usuarioCreacion, "")
    s.fechaPaseTecnico = Nz(rs!fechaPaseTecnico, 0)
    s.fechaCompletadoTecnico = Nz(rs!fechaCompletadoTecnico, 0)
    s.fechaModificacion = Nz(rs!fechaModificacion, 0)
    s.usuarioModificacion = Nz(rs!usuarioModificacion, "")
    
    Set MapRecordsetToSolicitud = s
End Function






