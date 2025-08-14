Attribute VB_Name = "modConfig"
' Modulo de Configuracion del Sistema CONDOR
' Gestiona la configuracion de rutas y entornos de desarrollo/produccion
' Version: 2.1 (corregido para compilacion)
' Fecha: 2025-01-14

Option Compare Database
Option Explicit

' Constante para modo de desarrollo
Public Const DEV_MODE As Boolean = True

' Constante para identificacion de aplicacion en sistema Lanzadera
Public Const IDAplicacion_CONDOR As Long = 231

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
Public Type T_AppConfig
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

' Variable global de configuracion
Public g_AppConfig As T_AppConfig

' Funcion principal de inicializacion del entorno
Public Function InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    Dim usarRutasLocales As Boolean
    
    ' Limpiar configuracion anterior
    Call ResetConfiguration
    
    ' --- LOGICA DE DECISION DE ENTORNO ---
    Select Case ENTORNO_FORZADO
        Case 1 ' ForzarLocal
            usarRutasLocales = True
            g_AppConfig.EntornoActivo = "Local (Forzado)"
        Case 2 ' ForzarRemoto
            usarRutasLocales = False
            g_AppConfig.EntornoActivo = "Remoto (Forzado)"
        Case 0 ' ForzarNinguno
            ' Comportamiento por defecto: depende del modo de compilacion
            usarRutasLocales = IsDevelopmentMode()
            If usarRutasLocales Then
                g_AppConfig.EntornoActivo = "Local (DEV_MODE)"
            Else
                g_AppConfig.EntornoActivo = "Remoto (Produccion)"
            End If
    End Select
    
    ' Configurar rutas segun la decision tomada
    If usarRutasLocales Then
        ' Configuracion para entorno de desarrollo LOCAL
        g_AppConfig.DatabasePath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
        g_AppConfig.dataPath = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
        g_AppConfig.ExpedientesPath = "C:\Proyectos\CONDOR\back\Expedientes_datos.accdb"
        g_AppConfig.PlantillasPath = "C:\Proyectos\CONDOR\back\recursos\Plantillas"
        g_AppConfig.LanzaderaDbPath = "C:\Proyectos\CONDOR\back\Lanzadera_Datos.accdb"
        g_AppConfig.sourcePath = "C:\Proyectos\CONDOR\src"
        g_AppConfig.backupPath = "C:\Proyectos\CONDOR\back\backups"
        g_AppConfig.logPath = "C:\Proyectos\CONDOR\logs"
        g_AppConfig.tempPath = "C:\Proyectos\CONDOR\temp"
    Else
        ' Configuracion para entorno de produccion/REMOTO
        g_AppConfig.DatabasePath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\CONDOR.accde"
        g_AppConfig.dataPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        g_AppConfig.ExpedientesPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\EXPEDIENTES_be.accdb"
        g_AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\Plantillas"
        g_AppConfig.LanzaderaDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
        g_AppConfig.sourcePath = "C:\Ruta\Invalida\En\Produccion"
        g_AppConfig.backupPath = Environ("APPDATA") & "\CONDOR\Backups"
        g_AppConfig.logPath = Environ("APPDATA") & "\CONDOR\Logs"
        g_AppConfig.tempPath = Environ("TEMP")
    End If
    
    ' Crear directorios si no existen
    Call CreateDirectoriesIfNeeded
    
    ' Marcar como inicializado
    g_AppConfig.IsInitialized = True
    
    InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    InitializeEnvironment = False
    g_AppConfig.IsInitialized = False
End Function

' Funcion para obtener el entorno activo (para debug)
Public Function GetActiveEnvironment() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetActiveEnvironment = g_AppConfig.EntornoActivo
End Function

' Funcion para obtener la ruta de la base de datos principal
Public Function GetDatabasePath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetDatabasePath = g_AppConfig.DatabasePath
End Function

' Funcion para obtener la ruta de la base de datos de datos
Public Function GetDataPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetDataPath = g_AppConfig.dataPath
End Function

' Funcion para obtener la ruta de la base de datos de Expedientes
Public Function GetExpedientesPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetExpedientesPath = g_AppConfig.ExpedientesPath
End Function

' Funcion para obtener la ruta de la base de datos de Expedientes (alias para compatibilidad)
Public Function GetExpedientesDbPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetExpedientesDbPath = g_AppConfig.ExpedientesPath
End Function

' Funcion para obtener la ruta de las plantillas
Public Function GetPlantillasPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetPlantillasPath = g_AppConfig.PlantillasPath
End Function

' Funcion para obtener la ruta de la base de datos Lanzadera
Public Function GetLanzaderaDbPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetLanzaderaDbPath = g_AppConfig.LanzaderaDbPath
End Function

' Funcion para obtener la ruta de codigo fuente
Public Function GetSourcePath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetSourcePath = g_AppConfig.sourcePath
End Function

' Funcion para obtener la ruta de backups
Public Function GetBackupPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetBackupPath = g_AppConfig.backupPath
End Function

' Funcion para obtener la ruta de logs
Public Function GetLogPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetLogPath = g_AppConfig.logPath
End Function

' Funcion para obtener la ruta temporal
Public Function GetTempPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetTempPath = g_AppConfig.tempPath
End Function

' Funcion privada para resetear la configuracion
Public Sub ResetConfiguration()
    g_AppConfig.DatabasePath = ""
    g_AppConfig.dataPath = ""
    g_AppConfig.ExpedientesPath = ""
    g_AppConfig.PlantillasPath = ""
    g_AppConfig.LanzaderaDbPath = ""
    g_AppConfig.sourcePath = ""
    g_AppConfig.backupPath = ""
    g_AppConfig.logPath = ""
    g_AppConfig.tempPath = ""
    g_AppConfig.IsInitialized = False
    g_AppConfig.EntornoActivo = ""
End Sub

' Funcion privada para crear directorios necesarios
Private Sub CreateDirectoriesIfNeeded()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear directorios locales si no existen
    If Not fso.FolderExists(g_AppConfig.backupPath) Then
        fso.CreateFolder g_AppConfig.backupPath
    End If
    
    If Not fso.FolderExists(g_AppConfig.logPath) Then
        fso.CreateFolder g_AppConfig.logPath
    End If
    
    If Not fso.FolderExists(g_AppConfig.tempPath) Then
        fso.CreateFolder g_AppConfig.tempPath
    End If
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

' Funcion para detectar si estamos en modo de desarrollo
Public Function IsDevelopmentMode() As Boolean
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

' Funcion de prueba
Public Function TestModConfig() As String
    Dim resultado As String
    resultado = "=== PRUEBA MODCONFIG ===" & vbCrLf
    
    ' Probar inicializacion
    Dim initResult As Boolean
    initResult = InitializeEnvironment()
    
    If initResult Then
        resultado = resultado & "[OK] InitializeEnvironment: OK" & vbCrLf
        resultado = resultado & "  |- Entorno Activo: " & GetActiveEnvironment() & vbCrLf
    Else
        resultado = resultado & "[FALLO] InitializeEnvironment: FALLO" & vbCrLf
    End If
    
    ' Probar obtencion de rutas
    Dim dbPath As String
    dbPath = GetDataPath()
    
    If Len(dbPath) > 0 Then
        resultado = resultado & "[OK] GetDataPath: OK -> " & dbPath & vbCrLf
    Else
        resultado = resultado & "[FALLO] GetDataPath: FALLO" & vbCrLf
    End If
    
    ' Probar GetLanzaderaDbPath
    Dim lanzaderaPath As String
    lanzaderaPath = GetLanzaderaDbPath()
    
    If Len(lanzaderaPath) > 0 Then
        resultado = resultado & "[OK] GetLanzaderaDbPath: OK -> " & lanzaderaPath & vbCrLf
    Else
        resultado = resultado & "[FALLO] GetLanzaderaDbPath: FALLO" & vbCrLf
    End If
    
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    TestModConfig = resultado
End Function