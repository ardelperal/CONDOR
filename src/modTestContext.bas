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
    If g_TestConfig Is Nothing Then
        Dim mockConfigImpl As New CMockConfig
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim testEnvPath As String
        testEnvPath = modTestUtils.GetProjectPath() & "back\test_env"
        
        ' COPIAS en workspace (solo estas 4 BDs)
        mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Lanzadera_workspace_test.accdb")
        mockConfigImpl.SetSetting "EXPEDIENTES_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Expedientes_integration_test.accdb")
        mockConfigImpl.SetSetting "CORREOS_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "correos_integration_test.accdb")
        mockConfigImpl.SetSetting "CONDOR_FRONT_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "CondorFront_integration_test.accdb")
        
        ' CONDOR BACK (directo) para el resto de dominios
        Dim condorBackPath As String
        Set g_TestConfig = mockConfigImpl ' Asignar temporalmente para usar la interfaz
        condorBackPath = g_TestConfig.GetValue("CONDOR_DATA_PATH")
        If Len(condorBackPath) = 0 Then condorBackPath = g_TestConfig.GetCondorDataPath()
        mockConfigImpl.SetSetting "CONDOR_DATA_PATH", condorBackPath
        mockConfigImpl.SetSetting "SOLICITUD_DATA_PATH", condorBackPath
        mockConfigImpl.SetSetting "DOCUMENT_DATA_PATH", condorBackPath
        mockConfigImpl.SetSetting "WORKFLOW_DATA_PATH", condorBackPath
        mockConfigImpl.SetSetting "MAPEO_DATA_PATH", condorBackPath
        mockConfigImpl.SetSetting "OPERATION_DATA_PATH", condorBackPath
        
        ' Passwords (si vacías, los repos NO deben añadir ;PWD=)
        mockConfigImpl.SetSetting "LANZADERA_PASSWORD", ""
        mockConfigImpl.SetSetting "EXPEDIENTES_PASSWORD", ""
        mockConfigImpl.SetSetting "CORREOS_PASSWORD", ""
        mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
        mockConfigImpl.SetSetting "CONDOR_FRONT_PASSWORD", ""
        
        ' Log en workspace
        mockConfigImpl.SetSetting "LOG_FILE_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "tests.log")
        
        ' Otras configuraciones que necesiten los tests
        mockConfigImpl.SetSetting "EMAIL_SENDER", "test@example.com"
        mockConfigImpl.SetSetting "EMAIL_SMTP_SERVER", "smtp.test.com"
        mockConfigImpl.SetSetting "EMAIL_SMTP_PORT", "587"
        mockConfigImpl.SetSetting "EMAIL_SMTP_USERNAME", "testuser"
        mockConfigImpl.SetSetting "EMAIL_SMTP_PASSWORD", "testpass"

        Set g_TestConfig = mockConfigImpl
        Set fso = Nothing
    End If
    Set GetTestConfig = g_TestConfig
End Function

' Procedimiento para forzar la destrucción del Singleton (útil entre ejecuciones de suites)
Public Sub ResetTestContext()
    Set g_TestConfig = Nothing
End Sub