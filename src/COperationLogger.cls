Option Compare Database
Option Explicit

' Clase: COperationLogger
' Descripción: Implementación concreta del servicio de logging de operaciones.
'              Registra las operaciones importantes en la tabla Tb_Operaciones_Log.

Implements IOperationLogger

Private m_configService As IConfig

Public Sub Initialize(config As IConfig)
    On Error GoTo ErrorHandler
    Set m_configService = config
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationLogger.Initialize")
End Sub



Private Sub IOperationLogger_LogOperation(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim logDbPath As String
    Dim logDbPassword As String
    
    ' Obtener la ruta y contraseña de la base de datos de logs de operaciones
    ' Asumimos que CConfig tiene métodos para esto o que se usa la DB principal
    logDbPath = m_configService.GetDataPath() ' Usamos DataPath como ejemplo, ajustar si hay una DB de logs dedicada
    logDbPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(logDbPath, False, False, ";PWD=" & logDbPassword)
    Set rs = db.OpenRecordset("Tb_Operaciones_Log", dbOpenDynaset)
    
    rs.AddNew
    rs("FechaHora") = Now
    rs("Usuario") = Environ("USERNAME") ' O el usuario autenticado de la aplicación
    rs("TipoOperacion") = operationType
    rs("IDEntidadAfectada") = entityId
    rs("Detalles") = details
    ' rs("VersionAplicacion") = modConfig.GetAppVersion() ' Descomentar si se implementa
    rs.Update
    
    rs.Close
    db.Close
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationLogger.IOperationLogger_LogOperation", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId)
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
End Sub
