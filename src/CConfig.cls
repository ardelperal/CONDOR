Option Compare Database
Option Explicit
' Clase: CConfig
' Descripci?n: Implementaci?n concreta del servicio de configuraci?n del sistema CONDOR
' Arquitectura: Capa de Servicios - Implementaci?n
' Implementa: IConfig

Implements IConfig

' Variables privadas de configuracion
Private m_Settings As Collection
Private m_OperationLogger As IOperationLogger
Private m_SolicitudRepository As ISolicitudRepository
Private m_IsInitialized As Boolean

' Método de inicialización para inyectar dependencias
Public Sub Initialize(operationLogger As IOperationLogger, solicitudRepository As ISolicitudRepository)
    On Error GoTo ErrorHandler
    
    Set m_OperationLogger = operationLogger
    Set m_SolicitudRepository = solicitudRepository
    Set m_Settings = New Collection
    m_IsInitialized = False
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.Initialize"
End Sub

' Implementaci?n de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    On Error GoTo ErrorHandler
    
    ' Logging de auditoría: Acceso a configuración
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "CONFIG_ACCESS", "Accediendo a configuración clave: " & key, "CConfig.GetValue"
    End If
    
    If Not m_IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    
    ' Búsqueda dinámica en la colección
    IConfig_GetValue = m_Settings(UCase(key))
    
    ' Logging de auditoría: Valor obtenido exitosamente
    If Not m_OperationLogger Is Nothing And Not IsNull(IConfig_GetValue) Then
        m_OperationLogger.LogOperation "CONFIG_VALUE_RETRIEVED", "Valor obtenido para clave: " & key, "CConfig.GetValue"
    End If
    
    Exit Function
    
ErrorHandler:
    If Not m_OperationLogger Is Nothing Then
        m_OperationLogger.LogOperation "CONFIG_ERROR", "Error accediendo configuración clave " & key & ": " & Err.Description, "CConfig.GetValue"
    End If
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetValue"
    IConfig_GetValue = Null
End Function

' Implementaci?n de IConfig.SetValue
Private Sub IConfig_SetValue(ByVal key As String, ByVal value As Variant)
    On Error GoTo ErrorHandler
    
    If m_Settings Is Nothing Then
        Set m_Settings = New Collection
    End If
    
    ' Remover clave existente si existe
    On Error Resume Next
    m_Settings.Remove UCase(key)
    On Error GoTo ErrorHandler
    
    ' Añadir nuevo valor
    m_Settings.Add value, UCase(key)
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.SetValue"
End Sub

Private Function IConfig_GetInstance() As IConfig
    Set IConfig_GetInstance = Me
End Function

' Implementaci?n de IConfig.InitializeEnvironment
Private Function IConfig_InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    Dim rs As Object
    
    ' Limpiar configuracion anterior
    Set m_Settings = New Collection
    
    ' Cargar configuración desde la base de datos usando el repositorio
    If m_SolicitudRepository Is Nothing Then
        modErrorHandler.LogError vbObjectError + 1001, "SolicitudRepository no inicializado", "CConfig.InitializeEnvironment"
        IConfig_InitializeEnvironment = False
        Exit Function
    End If
    
    ' Obtener recordset de configuración desde tbConfiguracion
    Set rs = m_SolicitudRepository.GetConfigurationSettings()
    
    If rs Is Nothing Then
        modErrorHandler.LogError vbObjectError + 1002, "No se pudo obtener configuración de la base de datos", "CConfig.InitializeEnvironment"
        IConfig_InitializeEnvironment = False
        Exit Function
    End If
    
    ' Cargar todos los pares clave/valor en la colección
    Do While Not rs.EOF
        m_Settings.Add rs.Fields("Valor").value, UCase(rs.Fields("Clave").value)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    ' Marcar como inicializado
    m_IsInitialized = True
    
    IConfig_InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.InitializeEnvironment"
    IConfig_InitializeEnvironment = False
    m_IsInitialized = False
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Function

' Implementaci?n de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    
    On Error Resume Next
    Dim testValue As Variant
    testValue = m_Settings(UCase(key))
    IConfig_HasKey = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.HasKey"
    IConfig_HasKey = False
End Function

' Implementaciones de la interfaz IConfig
Private Function IConfig_GetActiveEnvironment() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetActiveEnvironment = IConfig_GetValue("ENTORNOACTIVO")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetActiveEnvironment"
    IConfig_GetActiveEnvironment = ""
End Function

Private Function IConfig_GetDatabasePath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetDatabasePath = IConfig_GetValue("DATABASEPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetDatabasePath"
    IConfig_GetDatabasePath = ""
End Function

Private Function IConfig_GetDataPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetDataPath = IConfig_GetValue("DATAPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetDataPath"
    IConfig_GetDataPath = ""
End Function

Private Function IConfig_GetExpedientesPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetExpedientesPath = IConfig_GetValue("EXPEDIENTESPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetExpedientesPath"
    IConfig_GetExpedientesPath = ""
End Function

Private Function IConfig_GetExpedientesDbPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetExpedientesDbPath = IConfig_GetValue("EXPEDIENTESPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetExpedientesDbPath"
    IConfig_GetExpedientesDbPath = ""
End Function

Private Function IConfig_GetPlantillasPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetPlantillasPath = IConfig_GetValue("PLANTILLASPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetPlantillasPath"
    IConfig_GetPlantillasPath = ""
End Function

Private Function IConfig_GetLanzaderaDbPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetLanzaderaDbPath = IConfig_GetValue("LANZADERADBPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetLanzaderaDbPath"
    IConfig_GetLanzaderaDbPath = ""
End Function

Private Function IConfig_GetLogPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetLogPath = IConfig_GetValue("LOGPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetLogPath"
    IConfig_GetLogPath = ""
End Function

Private Function IConfig_GetDatabasePassword() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetDatabasePassword = IConfig_GetValue("DATABASEPASSWORD")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetDatabasePassword"
    IConfig_GetDatabasePassword = ""
End Function

Private Function IConfig_GetGeneratedDocsPath() As String
    On Error GoTo ErrorHandler
    
    IConfig_GetGeneratedDocsPath = IConfig_GetValue("GENERATEDDOCSPATH")
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.GetGeneratedDocsPath"
    IConfig_GetGeneratedDocsPath = ""
End Function

' Implementaci?n de IConfig.TestCConfig
Private Function IConfig_TestCConfig() As String
    On Error GoTo ErrorHandler
    
    Dim resultado As String
    resultado = "=== PRUEBA CCONFIG DINAMICA ===" & vbCrLf
    
    ' Probar inicializacion
    Dim initResult As Boolean
    initResult = IConfig_InitializeEnvironment()
    
    If initResult Then
        resultado = resultado & "[OK] InitializeEnvironment: OK" & vbCrLf
        resultado = resultado & "  |- Configuraciones cargadas: " & m_Settings.Count & vbCrLf
    Else
        resultado = resultado & "[FALLO] InitializeEnvironment: FALLO" & vbCrLf
    End If
    
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    IConfig_TestCConfig = resultado
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.TestCConfig"
    IConfig_TestCConfig = "Error en TestCConfig: " & Err.Description
End Function

Private Function IConfig_ValidateConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar que la configuración esté inicializada
    If Not m_IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    
    ' Validar que tengamos configuraciones cargadas
    If m_Settings.Count = 0 Then
        modErrorHandler.LogCriticalError vbObjectError + 1001, "No hay configuraciones cargadas", "CConfig.ValidateConfiguration"
        IConfig_ValidateConfiguration = False
        Exit Function
    End If
    
    IConfig_ValidateConfiguration = True
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.ValidateConfiguration"
    IConfig_ValidateConfiguration = False
End Function

Private Function IConfig_IsTestMode() As Boolean
    On Error GoTo ErrorHandler
    
    IConfig_IsTestMode = CBool(IConfig_GetValue("ISTESTMODE"))
    
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "CConfig.IsTestMode"
    IConfig_IsTestMode = False
End Function

'******************************************************************************
' MÉTODOS PÚBLICOS - DELEGACIÓN A LA INTERFAZ
'******************************************************************************

Public Function GetValue(ByVal key As String) As Variant
    GetValue = IConfig_GetValue(key)
End Function

Public Sub SetValue(ByVal key As String, ByVal value As Variant)
    IConfig_SetValue key, value
End Sub

Public Function GetActiveEnvironment() As String
    GetActiveEnvironment = IConfig_GetActiveEnvironment()
End Function

Public Function GetDatabasePath() As String
    GetDatabasePath = IConfig_GetDatabasePath()
End Function

Public Function GetDataPath() As String
    GetDataPath = IConfig_GetDataPath()
End Function

Public Function GetExpedientesPath() As String
    GetExpedientesPath = IConfig_GetExpedientesPath()
End Function

Public Function GetExpedientesDbPath() As String
    GetExpedientesDbPath = IConfig_GetExpedientesDbPath()
End Function

Public Function GetPlantillasPath() As String
    GetPlantillasPath = IConfig_GetPlantillasPath()
End Function

Public Function GetLanzaderaDbPath() As String
    GetLanzaderaDbPath = IConfig_GetLanzaderaDbPath()
End Function

Public Function GetLogPath() As String
    GetLogPath = IConfig_GetLogPath()
End Function

Public Function GetDatabasePassword() As String
    GetDatabasePassword = IConfig_GetDatabasePassword()
End Function

Public Function GetGeneratedDocsPath() As String
    GetGeneratedDocsPath = IConfig_GetGeneratedDocsPath()
End Function

Public Function TestCConfig() As String
    TestCConfig = IConfig_TestCConfig()
End Function

Public Function ValidateConfiguration() As Boolean
    ValidateConfiguration = IConfig_ValidateConfiguration()
End Function

Public Function IsTestMode() As Boolean
    IsTestMode = IConfig_IsTestMode()
End Function























