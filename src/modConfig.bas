Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit

' Modulo de Configuracion del Sistema CONDOR - Factory/Singleton
' Proporciona acceso a la instancia ?nica del servicio de configuraci?n
' Version: 3.0 (refactorizado para arquitectura de clases)
' Fecha: 2025-01-14


' Constante para modo de desarrollo
Public Const DEV_MODE As Boolean = True

' Constante para identificacion de aplicacion en sistema Lanzadera
Public Const IDAplicacion_CONDOR As Long = 231

' Instancia singleton del servicio de configuraci?n
Private g_ConfigInstance As IConfig

' Funci?n factory/singleton para obtener la instancia de configuraci?n
Public Function config() As IConfig
    If g_ConfigInstance Is Nothing Then
        Set g_ConfigInstance = New CConfig
    End If
    Set config = g_ConfigInstance
End Function

' Funci?n de inicializaci?n del entorno (delegada a CConfig)
Public Function InitializeEnvironment() As Boolean
    InitializeEnvironment = config().InitializeEnvironment()
End Function
' === FUNCIONES DE COMPATIBILIDAD (delegadas a CConfig) ===

' Funci?n para obtener el entorno activo (para debug)
Public Function GetActiveEnvironment() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetActiveEnvironment = configInstance.GetActiveEnvironment()
End Function

' Funci?n para obtener la ruta de la base de datos principal
Public Function GetDatabasePath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetDatabasePath = configInstance.GetDatabasePath()
End Function

' Funci?n para obtener la ruta de la base de datos de datos
Public Function GetDataPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetDataPath = configInstance.GetDataPath()
End Function

' Funci?n para obtener la ruta de la base de datos de Expedientes
Public Function GetExpedientesPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetExpedientesPath = configInstance.GetExpedientesPath()
End Function

' Funci?n para obtener la ruta de la base de datos de Expedientes (alias para compatibilidad)
Public Function GetExpedientesDbPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetExpedientesDbPath = configInstance.GetExpedientesDbPath()
End Function

' Funci?n para obtener la ruta de las plantillas
Public Function GetPlantillasPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetPlantillasPath = configInstance.GetPlantillasPath()
End Function

' Funci?n para obtener la ruta de la base de datos Lanzadera
Public Function GetLanzaderaDbPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetLanzaderaDbPath = configInstance.GetLanzaderaDbPath()
End Function

' Funci?n para obtener la ruta de c?digo fuente
Public Function GetSourcePath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetSourcePath = configInstance.GetSourcePath()
End Function

' Funci?n para obtener la ruta de backups
Public Function GetBackupPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetBackupPath = configInstance.GetBackupPath()
End Function

' Funci?n para obtener la ruta de logs
Public Function GetLogPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetLogPath = configInstance.GetLogPath()
End Function

' Funci?n para obtener la ruta temporal
Public Function GetTempPath() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    GetTempPath = configInstance.GetTempPath()
End Function

' Funci?n de prueba (delegada a CConfig)
Public Function TestModConfig() As String
    Dim configInstance As CConfig
    Set configInstance = config()
    TestModConfig = configInstance.TestCConfig()
End Function




