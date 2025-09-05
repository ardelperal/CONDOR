Attribute VB_Name = "modTestContext"
Option Compare Database
Option Explicit

' =================================================================
' MÓDULO: modTestContext
' DESCRIPCIÓN: Gestor Singleton para la configuración de pruebas.
' ARQUITECTURA: ÚNICA FUENTE DE VERDAD para la configuración en TODO el framework de testing.
' =================================================================

' Variable privada a nivel de módulo que contendrá la única instancia de la configuración.
Private g_TestConfig As IConfig

' FUNCIÓN SINGLETON: Crea la configuración en la primera llamada y la devuelve en las siguientes.
Public Function GetTestConfig() As IConfig
    ' Si la configuración no ha sido creada, créala y configúrala UNA SOLA VEZ.
    If g_TestConfig Is Nothing Then
        Dim mockConfigImpl As New CMockConfig
        Dim projectRoot As String
        projectRoot = modTestUtils.GetProjectPath()
        
        ' Configuración ESTÁNDAR para todas las bases de datos de prueba.
        mockConfigImpl.SetSetting "DATA_PATH", projectRoot & "back\test_db\active\CONDOR_integration_test.accdb"
        mockConfigImpl.SetSetting "DATABASE_PASSWORD", "" ' La BD principal de prueba no tiene pwd
        mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", projectRoot & "back\test_db\active\Lanzadera_integration_test.accdb"
        mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "EXPEDIENTES_DATA_PATH", projectRoot & "back\test_db\active\Expedientes_integration_test.accdb"
        mockConfigImpl.SetSetting "EXPEDIENTES_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "CORREOS_DB_PATH", projectRoot & "back\test_db\active\correos_integration_test.accdb"
        mockConfigImpl.SetSetting "CORREOS_PASSWORD", "dpddpd"
        
        ' Configuración general del entorno de prueba
        mockConfigImpl.SetSetting "LOG_FILE_PATH", projectRoot & "condor_test_run.log"
        mockConfigImpl.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
        mockConfigImpl.SetSetting "CORREO_ADMINISTRADOR", "admin.test@condor.com"
        
        Set g_TestConfig = mockConfigImpl
    End If
    
    ' Devuelve siempre la misma instancia.
    Set GetTestConfig = g_TestConfig
End Function

' Procedimiento para forzar la destrucción del Singleton (útil entre ejecuciones de suites)
Public Sub ResetTestContext()
    Set g_TestConfig = Nothing
End Sub