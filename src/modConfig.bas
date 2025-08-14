Attribute VB_Name = "modConfig"
' Módulo de Configuración del Sistema CONDOR
' Gestiona la configuración de rutas y entornos de desarrollo/producción
' Versión: 2.0 (con forzado de entorno para debug)
' Fecha: 2024

Option Compare Database
Option Explicit

#If DEV_MODE Then

' Constante para modo de desarrollo
Public Const DEV_MODE As Boolean = True

' Constante para identificación de aplicación en sistema Lanzadera
Public Const IDAplicacion_CONDOR As Long = 231

' ---------------------------------------------------------------------------------
' ==> INTERRUPTOR MANUAL PARA DESARROLLADORES <==
' Cambia este valor para forzar un entorno durante el desarrollo.
' ForzarLocal: Usa siempre las rutas locales (C:\Proyectos\...)
' ForzarRemoto: Usa siempre las rutas de red (\\datoste\...)
' ---------------------------------------------------------------------------------
Private Enum E_EnvironmentOverride
    ForzarNinguno = 0 ' Elige automáticamente basado en DEV_MODE
    ForzarLocal = 1
    ForzarRemoto = 2
End Enum
Private Const ENTORNO_FORZADO As E_EnvironmentOverride = ForzarLocal

' Estructura de configuración de la aplicación
Public Type T_AppConfig
    DatabasePath As String
    dataPath As String
    ExpedientesPath As String 'Añadido para la BBDD de Expedientes
    PlantillasPath As String 'Añadido para las plantillas Word
    LanzaderaDbPath As String 'Añadido para la BBDD de Lanzadera (gestión de roles)
    sourcePath As String
    backupPath As String
    logPath As String
    tempPath As String
    IsInitialized As Boolean
    EntornoActivo As String ' Para saber qué configuración se cargó
End Type

' Variable global de configuración
Public g_AppConfig As T_AppConfig

' Función principal de inicialización del entorno
Public Function InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    Dim usarRutasLocales As Boolean
    
    ' Limpiar configuración anterior
    Call ResetConfiguration
    
    ' --- LÓGICA DE DECISIÓN DE ENTORNO ---
    Select Case ENTORNO_FORZADO
        Case ForzarLocal
            usarRutasLocales = True
            g_AppConfig.EntornoActivo = "Local (Forzado)"
        Case ForzarRemoto
            usarRutasLocales = False
            g_AppConfig.EntornoActivo = "Remoto (Forzado)"
        Case ForzarNinguno
            ' Comportamiento por defecto: depende del modo de compilación
            usarRutasLocales = IsDevelopmentMode()
            If usarRutasLocales Then
                g_AppConfig.EntornoActivo = "Local (DEV_MODE)"
            Else
                g_AppConfig.EntornoActivo = "Remoto (Producción)"
            End If
    End Select
    
    ' Configurar rutas según la decisión tomada
    If usarRutasLocales Then
        ' Configuración para entorno de desarrollo LOCAL
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
        ' Configuración para entorno de producción/REMOTO
        ' OJO: La ruta de la BBDD de desarrollo (CONDOR.accdb) no debería estar en red,
        ' es el frontend que se distribuye. Aquí ponemos la ruta de la lanzadera.
        g_AppConfig.DatabasePath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\CONDOR.accde"
        g_AppConfig.dataPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        g_AppConfig.ExpedientesPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\EXPEDIENTES_be.accdb"
        g_AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\Plantillas"
        g_AppConfig.LanzaderaDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
        ' Estas rutas probablemente sigan siendo locales a la máquina del usuario
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

' Función para obtener el entorno activo (para debug)
Public Function GetActiveEnvironment() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetActiveEnvironment = g_AppConfig.EntornoActivo
End Function

' Función para obtener la ruta de la base de datos Lanzadera
Public Function GetLanzaderaDbPath() As String
    If Not g_AppConfig.IsInitialized Then
        Call InitializeEnvironment
    End If
    GetLanzaderaDbPath = g_AppConfig.LanzaderaDbPath
End Function

' ... (El resto de funciones Get...Path permanecen iguales) ...

' Función privada para resetear la configuración
Private Sub ResetConfiguration()
    ' ... (Sin cambios) ...
End Sub

' Función privada para crear directorios necesarios
Private Sub CreateDirectoriesIfNeeded()
    ' ... (Sin cambios) ...
End Sub

#End If

' Función de prueba (actualizada para mostrar el entorno)
Public Function TestModConfig() As String
    Dim resultado As String
    resultado = "=== PRUEBA MODCONFIG ===" & vbCrLf
    
    #If DEV_MODE Then
        ' Probar inicialización
        Dim initResult As Boolean
        initResult = InitializeEnvironment()
        
        If initResult Then
            resultado = resultado & "✓ InitializeEnvironment: OK" & vbCrLf
            resultado = resultado & "  └─ Entorno Activo: " & GetActiveEnvironment() & vbCrLf
        Else
            resultado = resultado & "✗ InitializeEnvironment: FALLO" & vbCrLf
        End If
        
        ' Probar obtención de ruta
        Dim dbPath As String
        dbPath = GetDataPath() 'Probamos con la de datos que es la que cambia
        
        If Len(dbPath) > 0 Then
            resultado = resultado & "✓ GetDataPath: OK -> " & dbPath & vbCrLf
        Else
            resultado = resultado & "✗ GetDataPath: FALLO" & vbCrLf
        End If
    #Else
        resultado = resultado & "✓ Modo producción - pruebas omitidas" & vbCrLf
    #End If
    
    resultado = resultado & "=== FIN PRUEBA ===" & vbCrLf
    TestModConfig = resultado
End Function
