Attribute VB_Name = "modTestContext"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO: modTestContext
' DESCRIPCIÓN: Gestiona el estado y contexto global para la ejecución de
'              la suite de pruebas (ej. Singletons, estado compartido).
' ============================================================================

' Variable a nivel de módulo para almacenar nuestra instancia Singleton
Private g_TestConfig As IConfig

' ============================================================================
' FUNCIÓN: GetTestConfig
' DESCRIPCIÓN: Devuelve una instancia única (Singleton) de IConfig para toda
'              la ejecución de la suite de pruebas. La carga solo ocurre una vez.
' RETORNA: IConfig - La instancia de configuración compartida.
' ============================================================================
Public Function GetTestConfig() As IConfig
    On Error GoTo ErrorHandler
    
    ' Si la configuración no ha sido cargada, créala y cárgala una sola vez.
    If g_TestConfig Is Nothing Then
        
        ' 1. Crear la instancia real de CConfig
        Dim configImpl As New CConfig
        
        ' 2. Crear un diccionario en memoria con la configuración estándar para pruebas
        Dim testSettings As New Scripting.Dictionary
        testSettings.CompareMode = TextCompare
        
        ' *** VALORES ESTÁNDAR PARA TODO EL ENTORNO DE PRUEBAS ***
        ' Usamos rutas relativas al proyecto para máxima portabilidad
        Dim basePath As String
        basePath = modTestUtils.GetProjectPath() ' Utilidad que ya tenemos
        
        testSettings.Add "DATA_PATH", basePath & "back\test_db\active\CONDOR_integration_test.accdb"
        testSettings.Add "DATABASE_PASSWORD", "" ' La BD de prueba no tiene contraseña
        testSettings.Add "LOG_FILE_PATH", basePath & "condor_test_run.log"
        testSettings.Add "USUARIO_ACTUAL", "test.user@condor.com"
        
        ' 3. Cargar la configuración en la instancia desde el diccionario
        configImpl.LoadFromDictionary testSettings
        
        ' 4. Asignar la instancia configurada a nuestra variable global (Singleton)
        Set g_TestConfig = configImpl
    End If
    
    ' 5. Devolver siempre la misma instancia
    Set GetTestConfig = g_TestConfig
    
    Exit Function

ErrorHandler:
    ' Si la configuración de pruebas falla, es un error fatal.
    Err.Raise Err.Number, "modTestContext.GetTestConfig", "Fallo crítico al inicializar la configuración de pruebas: " & Err.Description
End Function