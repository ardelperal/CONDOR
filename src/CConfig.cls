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
    dataPath As String
    ExpedientesPath As String 'Anadido para la BBDD de Expedientes
    PlantillasPath As String 'Anadido para las plantillas Word
    LanzaderaDbPath As String 'Anadido para la BBDD de Lanzadera (gestion de roles)
    sourcePath As String
    backupPath As String
    logPath As String
    tempPath As String
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
            IConfig_GetValue = m_AppConfig.dataPath
        Case "EXPEDIENTESPATH"
            IConfig_GetValue = m_AppConfig.ExpedientesPath
        Case "PLANTILLASPATH"
            IConfig_GetValue = m_AppConfig.PlantillasPath
        Case "LANZADERADBPATH"
            IConfig_GetValue = m_AppConfig.LanzaderaDbPath
        Case "SOURCEPATH"
            IConfig_GetValue = m_AppConfig.sourcePath
        Case "BACKUPPATH"
            IConfig_GetValue = m_AppConfig.backupPath
        Case "LOGPATH"
            IConfig_GetValue = m_AppConfig.logPath
        Case "TEMPPATH"
            IConfig_GetValue = m_AppConfig.tempPath
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
            m_AppConfig.dataPath = CStr(value)
        Case "EXPEDIENTESPATH"
            m_AppConfig.ExpedientesPath = CStr(value)
        Case "PLANTILLASPATH"
            m_AppConfig.PlantillasPath = CStr(value)
        Case "LANZADERADBPATH"
            m_AppConfig.LanzaderaDbPath = CStr(value)
        Case "SOURCEPATH"
            m_AppConfig.sourcePath = CStr(value)
        Case "BACKUPPATH"
            m_AppConfig.backupPath = CStr(value)
        Case "LOGPATH"
            m_AppConfig.logPath = CStr(value)
        Case "TEMPPATH"
            m_AppConfig.tempPath = CStr(value)
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
        m_AppConfig.DatabasePath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
        m_AppConfig.dataPath = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
        m_AppConfig.ExpedientesPath = "C:\Proyectos\CONDOR\back\Expedientes_datos.accdb"
        m_AppConfig.PlantillasPath = "C:\Proyectos\CONDOR\back\recursos\Plantillas"
        m_AppConfig.LanzaderaDbPath = "C:\Proyectos\CONDOR\back\Lanzadera_Datos.accdb"
        m_AppConfig.sourcePath = "C:\Proyectos\CONDOR\src"
        m_AppConfig.backupPath = "C:\Proyectos\CONDOR\back\backups"
        m_AppConfig.logPath = "C:\Proyectos\CONDOR\logs"
        m_AppConfig.tempPath = "C:\Proyectos\CONDOR\temp"
    Else
        ' Configuracion para entorno de produccion/REMOTO
        m_AppConfig.DatabasePath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\CONDOR.accde"
        m_AppConfig.dataPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        m_AppConfig.ExpedientesPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\EXPEDIENTES_be.accdb"
        m_AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\Plantillas"
        m_AppConfig.LanzaderaDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
        m_AppConfig.sourcePath = "C:\Ruta\Invalida\En\Produccion"
        m_AppConfig.backupPath = Environ("APPDATA") & "\CONDOR\Backups"
        m_AppConfig.logPath = Environ("APPDATA") & "\CONDOR\Logs"
        m_AppConfig.tempPath = Environ("TEMP")
    End If
    
    ' Crear directorios si no existen
    Call CreateDirectoriesIfNeeded
    
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
    IConfig_GetDataPath = m_AppConfig.dataPath
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

Private Function IConfig_GetSourcePath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetSourcePath = m_AppConfig.sourcePath
End Function

Private Function IConfig_GetBackupPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetBackupPath = m_AppConfig.backupPath
End Function

Private Function IConfig_GetLogPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If

    Dim logPath As String
    ' Asumimos que DEV_MODE es una constante o variable global accesible
    If DEV_MODE Then
        ' Entorno de Desarrollo: carpeta \logs en la raíz del proyecto
        logPath = Application.CurrentProject.Path & "\..\logs\"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
        IConfig_GetLogPath = logPath & "condor_development.log"
    Else
        ' Entorno de Producción: carpeta \logs junto al backend
        Dim backendPath As String
        backendPath = IConfig_GetDatabasePath()
        logPath = Left(backendPath, InStrRev(backendPath, "\")) & "logs\"
        If Dir(logPath, vbDirectory) = "" Then MkDir logPath
        IConfig_GetLogPath = logPath & "condor_production.log"
    End If
End Function

Private Function IConfig_GetTempPath() As String
    If Not m_AppConfig.IsInitialized Then
        Call IConfig_InitializeEnvironment
    End If
    IConfig_GetTempPath = m_AppConfig.tempPath
End Function

' M?todos privados de soporte
Private Sub ResetConfiguration()
    m_AppConfig.DatabasePath = ""
    m_AppConfig.dataPath = ""
    m_AppConfig.ExpedientesPath = ""
    m_AppConfig.PlantillasPath = ""
    m_AppConfig.LanzaderaDbPath = ""
    m_AppConfig.sourcePath = ""
    m_AppConfig.backupPath = ""
    m_AppConfig.logPath = ""
    m_AppConfig.tempPath = ""
    m_AppConfig.IsInitialized = False
    m_AppConfig.EntornoActivo = ""
End Sub

Private Sub CreateDirectoriesIfNeeded()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear directorios locales si no existen
    If Not fso.FolderExists(m_AppConfig.backupPath) Then
        fso.CreateFolder m_AppConfig.backupPath
    End If
    
    If Not fso.FolderExists(m_AppConfig.logPath) Then
        fso.CreateFolder m_AppConfig.logPath
    End If
    
    If Not fso.FolderExists(m_AppConfig.tempPath) Then
        fso.CreateFolder m_AppConfig.tempPath
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









