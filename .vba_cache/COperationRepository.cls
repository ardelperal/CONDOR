Option Compare Database
Option Explicit

' Clase: COperationRepository
' Descripción: Implementación concreta del repositorio de logs de operaciones.
'              Maneja la persistencia en la tabla Tb_Operaciones_Log.
' Arquitectura: Capa de Datos - Implementación del patrón Repository

Implements IOperationRepository

' Dependencia inyectada
Private m_configService As IConfig

' Método de inicialización para inyectar dependencias
Public Sub Initialize(config As IConfig)
    On Error GoTo ErrorHandler
    Set m_configService = config
    Exit Sub
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationRepository.Initialize")
End Sub

' Implementación del contrato IOperationRepository
Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim logDbPath As String
    Dim logDbPassword As String
    
    ' Obtener la ruta y contraseña de la base de datos de logs de operaciones
    logDbPath = m_configService.GetDataPath()
    logDbPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(logDbPath, False, False, ";PWD=" & logDbPassword)
    Set rs = db.OpenRecordset("Tb_Operaciones_Log", dbOpenDynaset)
    
    rs.AddNew
    rs("FechaHora") = Now
    rs("Usuario") = Environ("USERNAME") ' O el usuario autenticado de la aplicación
    rs("TipoOperacion") = operationType
    rs("IDEntidadAfectada") = entityId
    rs("Detalles") = details
    rs.Update
    
    rs.Close
    db.Close
    
    Exit Sub
    
ErrorHandler:
    Call modErrorHandler.LogError(Err.Number, Err.Description, "COperationRepository.IOperationRepository_SaveLog", _
                                  "TipoOperacion: " & operationType & ", Entidad: " & entityId)
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
End Sub