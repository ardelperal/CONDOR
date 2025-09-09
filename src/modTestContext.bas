Attribute VB_Name = "modTestContext"
Option Compare Database
Option Explicit

' =================================================================
' MÓDULO: modTestContext
' DESCRIPCIÓN: Gestor Singleton para la configuración de pruebas.
' Versión: 2.1 - Restaurado aprovisionamiento de BBDD (Lección 55)
' =================================================================

Private g_TestConfig As IConfig

Public Function GetTestConfig() As IConfig
    If g_TestConfig Is Nothing Then
        Dim mockConfigImpl As New CMockConfig
        
        ' --- RUTAS DE ORIGEN (MASTER FIXTURES EN /back) ---
        mockConfigImpl.SetSetting "DEV_CONDOR_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\CONDOR_datos.accdb")
        mockConfigImpl.SetSetting "DEV_LANZADERA_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\Lanzadera_datos.accdb")
        mockConfigImpl.SetSetting "DEV_EXPEDIENTES_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\Expedientes_datos.accdb")
        mockConfigImpl.SetSetting "DEV_CORREOS_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\correos_datos.accdb")
        mockConfigImpl.SetSetting "PRODUCTION_TEMPLATES_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\recursos\Plantillas\")

        ' --- RUTAS DE DESTINO (COPIAS DE TRABAJO EN /workspace) ---
        Dim testTemplatesPath As String: testTemplatesPath = modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "test_templates\")
        mockConfigImpl.SetSetting "TEMPLATES_PATH", testTemplatesPath
        mockConfigImpl.SetSetting "GENERATED_DOCS_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "generated_documents\")
        mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "CONDOR_integration_test.accdb")
        mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Lanzadera_integration_test.accdb")
        mockConfigImpl.SetSetting "EXPEDIENTES_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Expedientes_integration_test.accdb")
        mockConfigImpl.SetSetting "CORREOS_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Correos_integration_test.accdb")

        ' --- NOMBRES DE FICHEROS DE PLANTILLAS (sin cambios) ---
        mockConfigImpl.SetSetting "TEMPLATE_PC_FILENAME", "F4203.11 - Propuesta de Cambio.docx"
        mockConfigImpl.SetSetting "TEMPLATE_CDCA_FILENAME", "F4203.10 - Desviacion Concesion.docx"
        mockConfigImpl.SetSetting "TEMPLATE_CDCASUB_FILENAME", "F4203.101 - Desviacion Concesion Sub-suministrador.docx"
        
        ' --- OTRAS CONFIGURACIONES DE PRUEBA (sin cambios) ---
        mockConfigImpl.SetSetting "LOG_FILE_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "tests.log")
        mockConfigImpl.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
        mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", "231"
        mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "EXPEDIENTES_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "CORREOS_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
        
        Set g_TestConfig = mockConfigImpl
    End If
    Set GetTestConfig = g_TestConfig
End Function

Public Sub ResetTestContext()
    Set g_TestConfig = Nothing
End Sub