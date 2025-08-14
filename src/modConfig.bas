Attribute VB_Name "modConfig"
' Definir constante de compilacion para desarrollo
#Const DEV_MODE = True

' =============================================================================
' MODULO DE CONFIGURACION Y GESTION DE ENTORNOS
' =============================================================================
' Proposito: Gestionar las rutas y configuraciones segun el entorno
' Entornos: Local (desarrollo) y Remoto (produccion)
' =============================================================================

' Estructura para almacenar todas las rutas de configuracion
Private Type T_AppConfig
    CondorDbPath As String
    ExpedientesDbPath As String
    PlantillasPath As String
   
End Type

' Variable privada a nivel de modulo para guardar la configuracion
Private AppConfig As T_AppConfig
Private ConfigInitialized As Boolean

' =============================================================================
' FUNCION PRINCIPAL DE INICIALIZACION
' =============================================================================

' Inicializa el entorno segun el modo de compilacion
Public Sub InitializeEnvironment()
    
    #If DEV_MODE Then
        ' ENTORNO DE DESARROLLO - Rutas locales
        AppConfig.CondorDbPath = "./back/CONDOR_datos.accdb"
        AppConfig.ExpedientesDbPath = "./back/Expedientes_datos.accdb""
        AppConfig.PlantillasPath = "./back/recursos/Plantillas/"
        
    #Else
        ' ENTORNO DE PRODUCCION - Rutas de red
        AppConfig.CondorDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\CONDOR_datos.accdb"
        AppConfig.ExpedientesDbPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb"
        AppConfig.PlantillasPath = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR Prueba\recursos\Plantillas\"
       
    
    
    ConfigInitialized = True
    
End Sub

' =============================================================================
' FUNCIONES DE ACCESO A RUTAS
' =============================================================================

' Obtiene la ruta de la base de datos de CONDOR
Public Function GetCondorDbPath() As String
    If Not ConfigInitialized Then InitializeEnvironment
    GetCondorDbPath = AppConfig.CondorDbPath
End Function

' Obtiene la ruta de la base de datos de Expedientes
Public Function GetExpedientesDbPath() As String
    If Not ConfigInitialized Then InitializeEnvironment
    GetExpedientesDbPath = AppConfig.ExpedientesDbPath
End Function

' Obtiene la ruta de las plantillas de documentos
Public Function GetPlantillasPath() As String
    If Not ConfigInitialized Then InitializeEnvironment
    GetPlantillasPath = AppConfig.PlantillasPath
End Function



' =============================================================================
' FUNCIONES DE UTILIDAD
' =============================================================================

' Obtiene informacion sobre el entorno actual
Public Function GetEnvironmentInfo() As String
    If Not ConfigInitialized Then InitializeEnvironment
    
    #If DEV_MODE Then
        GetEnvironmentInfo = "DESARROLLO (Local)"
    #Else
        GetEnvironmentInfo = "PRODUCCION (Remoto)"
    #End If
End Function

' Verifica si la configuracion ha sido inicializada
Public Function IsConfigInitialized() As Boolean
    IsConfigInitialized = ConfigInitialized
End Function

' Fuerza la reinicializacion de la configuracion
Public Sub ResetConfiguration()
    ConfigInitialized = False
    InitializeEnvironment
End Sub

#End If