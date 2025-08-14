Attribute VB_Name = "modConfig"
' Modulo de Configuracion del Sistema CONDOR
' Gestiona la configuracion de rutas y entornos de desarrollo/produccion
' Version: 2.0 (con forzado de entorno para debug)
' Fecha: 2024

Option Compare Database
Option Explicit

#If DEV_MODE Then

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
        g_AppConfig.ExpedientesPath = "C:\Proyectos\CONDOR\back\EXPEDIENTES.accdb" 'Ruta local de ejemplo
        g_AppConfig.PlantillasPath = "C:\Proyectos\CONDOR\templates"
        g_AppConfig.LanzaderaDbPath = "C:\Proyectos\CONDOR\back\Lanzadera_Datos.accdb"
        g_AppConfig.sourcePath = "C:\Proyectos\CONDOR\src"
        g_AppConfig.backupPath = "C:\Proyectos\CONDOR\back\backups"
        g_AppConfig.logPath = "C:\Proyectos\CONDOR\logs"
        g_AppConfig.tempPath = "C:\Proyectos\CONDOR\temp"
    Else
        ' Configuracion para entorno de produccion/REMOTO
        ' OJO: La ruta de la BBDD de desarrollo (CONDOR.accdb) no deberia estar en red,
        ' es el frontend que se distribuye. Aqui ponemos la ruta de la lanzadera.
        g_AppConfig.DatabasePath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\CONDOR.accde"
        g_AppConfig.dataPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        g_AppConfig.ExpedientesPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\EXPEDIENTES_be.accdb"
        g_AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\Plantillas"
        g_AppConfig.LanzaderaDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
        ' Estas rutas probablemente sigan siendo locales a la maquina del usuario
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

' Funcion para obtener la ruta de la base de datos Lanzadera
Public Function GetLanzaderaDbPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetLanzaderaDbPath = g_AppConfig.LanzaderaDbPath
End Function

' ... (El resto de funciones Get...Path permanecen iguales) ...

' Funcion privada para resetear la configuracion
Private Sub ResetConfiguration()
    ' ... (Sin cambios) ...
End Sub

' Funcion privada para crear directorios necesarios
Private Sub CreateDirectoriesIfNeeded()
    ' ... (Sin cambios) ...
End Sub

#End If

' Funcion para detectar si estamos en modo de desarrollo
' Detecta basandose en la existencia de la estructura de directorios de desarrollo
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
            Exit Function
        End If
    Next i
    
    ' Si no encuentra ninguna ruta de desarrollo, asumir produccion
    IsDevelopmentMode = False
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir modo produccion por seguridad
    IsDevelopmentMode = False
End Function

' Funcion de prueba (actualizada para mostrar el entorno)
Public Function TestModConfig() As String
    Dim resultado As String
    resultado = "=== PRUEBA MODCONFIG ===" & vbCrLf
    
    #If DEV_MODE Then
        ' Probar inicializacion
        Dim initResult As Boolean
        initResult = InitializeEnvironment()
        
        If initResult Then
            resultado = resultado & "✓ InitializeEnvironment: OK" & vbCrLf
            resultado = resultado & "  └─ Entorno Activo: " & GetActiveEnvironment() & vbCrLf
        Else
            resultado = resultado & "✗ InitializeEnvironment: FALLO" & vbCrLf
        End If
        
        ' Probar obtencion de ruta
        Dim dbPath As String
        dbPath = GetDataPath() 'Probamos con la de datos que es la que cambia
        
        If Len(dbPath) > 0 Then
            resultado = resultado & "✓ GetDataPath: OK -> " & dbPath & vbCrLf
        Else
            resultado = resultado & "✗ GetDataPath: FALLO" & vbCrLf
        End If
    #Else
        resultado = resultado & "Modo produccion - pruebas omitidas" & vbCrLf
    #End If
    
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    TestModConfig = resultado
End Function