Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit
' Factory para el servicio de configuración. Versión 4.0.



Public Function CreateConfigService() As IConfig
    On Error GoTo ErrorHandler
    
    Dim config As New CConfig
    Dim dbFrontend As DAO.Database
    Dim dbBackend As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendPath As String
    Dim backendPassword As String
    
    ' Conectarse a la base de datos del Frontend (CurrentDb)
    Set dbFrontend = CurrentDb
    
    ' Leer configuración local desde TbLocalConfig
    Set rs = dbFrontend.OpenRecordset("SELECT BackendPath, BackendPassword FROM TbLocalConfig", dbOpenSnapshot)
    
    If Not rs.EOF Then
        backendPath = rs.Fields("BackendPath").Value
        backendPassword = rs.Fields("BackendPassword").Value
    Else
        ' Valores por defecto si no existe configuración local
        backendPath = Application.CurrentProject.Path & "\..\data\CONDOR_Backend.accdb"
        backendPassword = "dpddpd"
    End If
    
    rs.Close
    Set rs = Nothing
    
    ' Conectarse a la base de datos del Backend
    Set dbBackend = DBEngine.OpenDatabase(backendPath, False, False, "MS Access;PWD=" & backendPassword)
    
    ' Cargar configuración desde el Backend
    config.Load dbBackend
    
    ' Cerrar conexión al Backend
    dbBackend.Close
    Set dbBackend = Nothing
    
    ' Devolver instancia configurada
    Set CreateConfigService = config
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not dbBackend Is Nothing Then
        dbBackend.Close
        Set dbBackend = Nothing
    End If
    
    ' En caso de error, devolver instancia vacía
    Set CreateConfigService = config
End Function











