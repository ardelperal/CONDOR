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
    Dim entorno As String
    
    ' Constantes con las rutas base según el entorno
    Const BASE_PATH_LOCAL As String = "C:\Proyectos\CONDOR\"
    Const BASE_PATH_OFICINA As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR\"
    
    ' PASO 1: Conectarse al Frontend local para leer configuración de arranque
    Set dbFrontend = CurrentDb
    
    ' PASO 2: Leer el entorno configurado desde TbLocalConfig del Frontend
    Set rs = dbFrontend.OpenRecordset("SELECT Entorno FROM TbLocalConfig", dbOpenSnapshot)
    
    If Not rs.EOF Then
        entorno = rs.Fields("Entorno").Value
    Else
        Err.Raise vbObjectError + 1001, "CreateConfigService", "No se encontró configuración de entorno en TbLocalConfig"
    End If
    
    rs.Close
    Set rs = Nothing
    
    ' PASO 3: Construir ruta al Backend usando valor de Entorno
    Select Case UCase(entorno)
        Case "LOCAL"
            backendPath = BASE_PATH_LOCAL & "back\CONDOR_datos.accdb"
        Case "OFICINA"
            backendPath = BASE_PATH_OFICINA & "back\CONDOR_datos.accdb"
        Case Else
            Err.Raise vbObjectError + 1002, "CreateConfigService", "Entorno no válido: '" & entorno & "'. Los valores válidos son 'LOCAL' o 'OFICINA'"
    End Select
    
    ' Contraseña fija para el backend
    backendPassword = "dpddpd"
    
    ' PASO 4: Conectarse al Backend con la ruta correcta
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
    If Not dbFrontend Is Nothing Then
        dbFrontend.Close
        Set dbFrontend = Nothing
    End If
    If Not dbBackend Is Nothing Then
        dbBackend.Close
        Set dbBackend = Nothing
    End If
    
    ' En caso de error, devolver instancia vacía
    Set CreateConfigService = config
End Function











