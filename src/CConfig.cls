VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

' Clase: CConfig
' Descripci?n: Implementaci?n concreta del servicio de configuraci?n del sistema CONDOR
' Arquitectura: Capa de Servicios - Implementaci?n
' Implementa: IConfig

Implements IConfig

' Constante para identificacion de aplicacion en sistema Lanzadera
' (Constante movida a modConfig.bas para cumplir con restricciones VBA)

' ---------------------------------------------------------------------------------
' ==> INTERRUPTOR MANUAL PARA DESARROLLADORES <==
' Cambia este valor para forzar un entorno durante el desarrollo.
' ForzarLocal: Usa siempre las rutas locales (C:\Proyectos\...)
' ForzarRemoto: Usa siempre las rutas de red (\\datoste\...)
' ---------------------------------------------------------------------------------
Private Enum E_EnvironmentOverride
    ForzarNinguno = 0 ' Elige automaticamente basado en DEV_MODE
    ForzarLocal = 1
    ForzarRemoto = 2
End Enum
Private Const ENTORNO_FORZADO As Long = 1 ' ForzarLocal

' Estructura de configuracion de la aplicacion
Private Type T_AppConfig
    DatabasePath As String
    DataPath As String
    ExpedientesPath As String 'Anadido para la BBDD de Expedientes
    PlantillasPath As String 'Anadido para las plantillas Word
    plantillaPCPath As String 'Ruta completa a PC.docx
    plantillaCDCAPath As String 'Ruta completa a CD_CA.docx
    plantillaCDCASUBPath As String 'Ruta completa a CD_CA_SUB.docx
    LanzaderaDbPath As String 'Anadido para la BBDD de Lanzadera (gestion de roles)
    LogPath As String
    DatabasePassword As String 'Contraseña para bases de datos back
    IsInitialized As Boolean
    EntornoActivo As String ' Para saber que configuracion se cargo
End Type

' Variable privada de configuracion
Private m_AppConfig As T_AppConfig

' Implementaci?n de IConfig.GetValue
Private Function IConfig_GetValue(ByVal key As String) As Variant
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    
    Select Case UCase(key)
        Case "DATABASEPATH"
            IConfig_GetValue = m_AppConfig.DatabasePath
        Case "DATAPATH"
            IConfig_GetValue = m_AppConfig.DataPath
        Case "EXPEDIENTESPATH"
            IConfig_GetValue = m_AppConfig.ExpedientesPath
        Case "PLANTILLASPATH"
            IConfig_GetValue = m_AppConfig.PlantillasPath
        Case "PLANTILLAPCPATH"
            IConfig_GetValue = m_AppConfig.plantillaPCPath
        Case "PLANTILLACDCAPATH"
            IConfig_GetValue = m_AppConfig.plantillaCDCAPath
        Case "PLANTILLACDCASUBPATH"
            IConfig_GetValue = m_AppConfig.plantillaCDCASUBPath
        Case "LANZADERADBPATH"
            IConfig_GetValue = m_AppConfig.LanzaderaDbPath
        Case "LOGPATH"
            IConfig_GetValue = m_AppConfig.LogPath
        Case "DATABASEPASSWORD"
            IConfig_GetValue = m_AppConfig.DatabasePassword
        Case "ENTORNOACTIVO"
            IConfig_GetValue = m_AppConfig.EntornoActivo
        Case "ISINITIALIZED"
            IConfig_GetValue = m_AppConfig.IsInitialized
        Case "IDAPLICACION_CONDOR"
            IConfig_GetValue = IDAplicacion_CONDOR
        Case Else
            IConfig_GetValue = Null
    End Select
End Function

' Implementaci?n de IConfig.SetValue
Private Sub IConfig_SetValue(ByVal key As String, ByVal value As Variant)
    Select Case UCase(key)
        Case "DATABASEPATH"
            m_AppConfig.DatabasePath = CStr(value)
        Case "DATAPATH"
            m_AppConfig.DataPath = CStr(value)
        Case "EXPEDIENTESPATH"
            m_AppConfig.ExpedientesPath = CStr(value)
        Case "PLANTILLASPATH"
            m_AppConfig.PlantillasPath = CStr(value)
        Case "PLANTILLAPCPATH"
            m_AppConfig.plantillaPCPath = CStr(value)
        Case "PLANTILLACDCAPATH"
            m_AppConfig.plantillaCDCAPath = CStr(value)
        Case "PLANTILLACDCASUBPATH"
            m_AppConfig.plantillaCDCASUBPath = CStr(value)
        Case "LANZADERADBPATH"
            m_AppConfig.LanzaderaDbPath = CStr(value)
        Case "LOGPATH"
            m_AppConfig.LogPath = CStr(value)
        Case "DATABASEPASSWORD"
            m_AppConfig.DatabasePassword = CStr(value)
        Case "ENTORNOACTIVO"
            m_AppConfig.EntornoActivo = CStr(value)
        Case "ISINITIALIZED"
            m_AppConfig.IsInitialized = CBool(value)
    End Select
End Sub

' Implementaci?n de IConfig.InitializeEnvironment
Private Function IConfig_InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    Dim usarRutasLocales As Boolean
    
    ' Limpiar configuracion anterior
    Call ResetConfiguration
    
    ' --- LOGICA DE DECISION DE ENTORNO ---
    Select Case ENTORNO_FORZADO
        Case 1 ' ForzarLocal
            usarRutasLocales = True
            m_AppConfig.EntornoActivo = "Local (Forzado)"
        Case 2 ' ForzarRemoto
            usarRutasLocales = False
            m_AppConfig.EntornoActivo = "Remoto (Forzado)"
        Case 0 ' ForzarNinguno
            ' Comportamiento por defecto: depende del modo de compilacion
            usarRutasLocales = IsDevelopmentMode()
            If usarRutasLocales Then
                m_AppConfig.EntornoActivo = "Local (DEV_MODE)"
            Else
                m_AppConfig.EntornoActivo = "Remoto (Produccion)"
            End If
    End Select
    
    ' Configurar rutas segun la decision tomada
    If usarRutasLocales Then
        ' Configuracion para entorno de desarrollo LOCAL
        Dim basePath As String
        basePath = CurrentProject.Path & "\.."
        m_AppConfig.DatabasePath = basePath & "back\Desarrollo\CONDOR.accdb"
        m_AppConfig.DataPath = basePath & "back\CONDOR_datos.accdb"
        m_AppConfig.ExpedientesPath = basePath & "back\Expedientes_datos.accdb"
        m_AppConfig.PlantillasPath = basePath & "back\recursos\Plantillas"
        m_AppConfig.LanzaderaDbPath = basePath & "back\Lanzadera_Datos.accdb"
        m_AppConfig.LogPath = basePath & "logs"
        m_AppConfig.DatabasePassword = "dpddpd"
    Else
        ' Configuracion para entorno de produccion/REMOTO
        m_AppConfig.DatabasePath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\CONDOR.accde"
        m_AppConfig.DataPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        m_AppConfig.ExpedientesPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\EXPEDIENTES_be.accdb"
        m_AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\Plantillas"
        m_AppConfig.LanzaderaDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
        m_AppConfig.LogPath = Environ("APPDATA") & "\CONDOR\Logs"
        m_AppConfig.DatabasePassword = "dpddpd"
    End If
    
    ' Construir rutas completas a las plantillas
    m_AppConfig.plantillaPCPath = m_AppConfig.PlantillasPath & "\PC.docx"
    m_AppConfig.plantillaCDCAPath = m_AppConfig.PlantillasPath & "\CD_CA.docx"
    m_AppConfig.plantillaCDCASUBPath = m_AppConfig.PlantillasPath & "\CD_CA_SUB.docx"
    
    ' Crear directorios si no existen
    Call CreateDirectoriesIfNeeded
    
    ' Validar la configuración cargada
    If Not ValidateConfiguration() Then
        m_AppConfig.IsInitialized = False
        IConfig_InitializeEnvironment = False
        Exit Function
    End If

    ' Marcar como inicializado
    m_AppConfig.IsInitialized = True
    
    IConfig_InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    IConfig_InitializeEnvironment = False
    m_AppConfig.IsInitialized = False
End Function

' Implementaci?n de IConfig.HasKey
Private Function IConfig_HasKey(ByVal key As String) As Boolean
    Select Case UCase(key)
        Case "DATABASEPATH", "DATAPATH", "EXPEDIENTESPATH", "PLANTILLASPATH", _
             "LANZADERADBPATH", "SOURCEPATH", "BACKUPPATH", "LOGPATH", _
             "TEMPPATH", "ENTORNOACTIVO", "ISINITIALIZED", "IDAPLICACION_CONDOR"
            IConfig_HasKey = True
        Case Else
            IConfig_HasKey = False
    End Select
End Function

' Implementaciones de la interfaz IConfig
Private Function IConfig_GetActiveEnvironment() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetActiveEnvironment = m_AppConfig.EntornoActivo
End Function

Private Function IConfig_GetDatabasePath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetDatabasePath = m_AppConfig.DatabasePath
End Function

Private Function IConfig_GetDataPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetDataPath = m_AppConfig.DataPath
End Function

Private Function IConfig_GetExpedientesPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetExpedientesPath = m_AppConfig.ExpedientesPath
End Function

Private Function IConfig_GetExpedientesDbPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetExpedientesDbPath = m_AppConfig.ExpedientesPath
End Function

Private Function IConfig_GetPlantillasPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetPlantillasPath = m_AppConfig.PlantillasPath
End Function

Private Function IConfig_GetLanzaderaDbPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetLanzaderaDbPath = m_AppConfig.LanzaderaDbPath
End Function

Private Function IConfig_GetLogPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If

    Dim LogPath As String
    ' Asumimos que DEV_MODE es una constante o variable global accesible
    If DEV_MODE Then
        ' Entorno de Desarrollo: carpeta \logs en la raíz del proyecto
        LogPath = Application.CurrentProject.path & "\..\logs\"
        If Dir(LogPath, vbDirectory) = "" Then MkDir LogPath
        IConfig_GetLogPath = LogPath & "condor_development.log"
    Else
        ' Entorno de Producción: carpeta \logs junto al backend
        Dim backendPath As String
        backendPath = IConfig_GetDatabasePath()
        LogPath = Left(backendPath, InStrRev(backendPath, "\")) & "logs\"
        If Dir(LogPath, vbDirectory) = "" Then MkDir LogPath
        IConfig_GetLogPath = LogPath & "condor_production.log"
    End If
End Function

Private Function IConfig_GetDatabasePassword() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetDatabasePassword = m_AppConfig.DatabasePassword
End Function

' M?todos privados de soporte
Private Sub ResetConfiguration()
    m_AppConfig.DatabasePath = ""
    m_AppConfig.DataPath = ""
    m_AppConfig.ExpedientesPath = ""
    m_AppConfig.PlantillasPath = ""
    m_AppConfig.plantillaPCPath = ""
    m_AppConfig.plantillaCDCAPath = ""
    m_AppConfig.plantillaCDCASUBPath = ""
    m_AppConfig.LanzaderaDbPath = ""
    m_AppConfig.LogPath = ""
    m_AppConfig.DatabasePassword = ""
    m_AppConfig.IsInitialized = False
    m_AppConfig.EntornoActivo = ""
End Sub

Private Sub CreateDirectoriesIfNeeded()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear directorios locales si no existen
    If Not fso.FolderExists(m_AppConfig.LogPath) Then
        fso.CreateFolder m_AppConfig.LogPath
    End If
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Function IsDevelopmentMode() As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar si existe la estructura de directorios de desarrollo
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar rutas caracteristicas del entorno de desarrollo
    Dim devPaths As Variant
    devPaths = Array( _
        "C:\Proyectos\CONDOR\src", _
        "C:\Proyectos\CONDOR\back\Desarrollo", _
        "C:\Proyectos\CONDOR\README.md" _
    )
    
    Dim i As Integer
    For i = 0 To UBound(devPaths)
        If fso.FolderExists(devPaths(i)) Or fso.FileExists(devPaths(i)) Then
            IsDevelopmentMode = True
            Set fso = Nothing
            Exit Function
        End If
    Next i
    
    ' Si no encuentra ninguna ruta de desarrollo, asumir produccion
    IsDevelopmentMode = False
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir modo produccion por seguridad
    IsDevelopmentMode = False
    If Not fso Is Nothing Then Set fso = Nothing
End Function

' Implementaci?n de IConfig.TestCConfig
Private Function IConfig_TestCConfig() As String
    Dim resultado As String
    resultado = "=== PRUEBA CCONFIG ===" & vbCrLf
    
    ' Probar inicializacion
    Dim initResult As Boolean
    initResult = IConfig_InitializeEnvironment()
    
    If initResult Then
        resultado = resultado & "[OK] InitializeEnvironment: OK" & vbCrLf
        resultado = resultado & "  |- Entorno Activo: " & IConfig_GetActiveEnvironment() & vbCrLf
    Else
        resultado = resultado & "[FALLO] InitializeEnvironment: FALLO" & vbCrLf
    End If
    
    ' Probar obtencion de rutas
    Dim dbPath As String
    dbPath = IConfig_GetDataPath()
    
    If Len(dbPath) > 0 Then
        resultado = resultado & "[OK] GetDataPath: OK -> " & dbPath & vbCrLf
    Else
        resultado = resultado & "[FALLO] GetDataPath: FALLO" & vbCrLf
    End If
    
    ' Probar GetLanzaderaDbPath
    Dim lanzaderaPath As String
    lanzaderaPath = IConfig_GetLanzaderaDbPath()
    
    If Len(lanzaderaPath) > 0 Then
        resultado = resultado & "[OK] GetLanzaderaDbPath: OK -> " & lanzaderaPath & vbCrLf
    Else
        resultado = resultado & "[FALLO] GetLanzaderaDbPath: FALLO" & vbCrLf
    End If
    
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    IConfig_TestCConfig = resultado
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

Public Function InitializeEnvironment() As Boolean
    InitializeEnvironment = IConfig_InitializeEnvironment()
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = IConfig_HasKey(key)
End Function

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

Public Function TestCConfig() As String
    TestCConfig = IConfig_TestCConfig()
End Function

Private Function ValidateConfiguration() As Boolean
    ' Comprueba que las rutas esenciales han sido cargadas
    ValidateConfiguration = True ' Asumir éxito por defecto

    If IsNull(m_AppConfig.DatabasePath) Or m_AppConfig.DatabasePath = "" Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1001, "Error de configuración: La ruta 'DatabasePath' no está definida.", "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Exit Function
    End If
    
    If IsNull(m_AppConfig.LanzaderaDbPath) Or m_AppConfig.LanzaderaDbPath = "" Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1002, "Error de configuración: La ruta 'LanzaderaDbPath' no está definida.", "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Exit Function
    End If
    
    If IsNull(m_AppConfig.ExpedientesPath) Or m_AppConfig.ExpedientesPath = "" Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1003, "Error de configuración: La ruta 'ExpedientesPath' no está definida.", "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Validar existencia de plantillas
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(m_AppConfig.plantillaPCPath) Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1004, "Error de configuración: La plantilla PC.docx no existe en la ruta: " & m_AppConfig.plantillaPCPath, "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Set fso = Nothing
        Exit Function
    End If
    
    If Not fso.FileExists(m_AppConfig.plantillaCDCAPath) Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1005, "Error de configuración: La plantilla CD_CA.docx no existe en la ruta: " & m_AppConfig.plantillaCDCAPath, "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Set fso = Nothing
        Exit Function
    End If
    
    If Not fso.FileExists(m_AppConfig.plantillaCDCASUBPath) Then
        Call modErrorHandler.LogCriticalError(vbObjectError + 1006, "Error de configuración: La plantilla CD_CA_SUB.docx no existe en la ruta: " & m_AppConfig.plantillaCDCASUBPath, "CConfig.ValidateConfiguration")
        ValidateConfiguration = False
        Set fso = Nothing
        Exit Function
    End If
    
    Set fso = Nothing
    
End Function





















