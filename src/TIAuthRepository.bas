Attribute VB_Name = "TIAuthRepository"
Option Compare Database
Option Explicit


' ============================================================================
' MÓDULO DE PRUEBAS DE INTEGRACIÓN PARA CAuthRepository
' ============================================================================
' Este módulo contiene pruebas de integración para CAuthRepository
' que validan la conexión y consultas contra la base de datos real de la Lanzadera.

' Constantes para el autoaprovisionamiento de bases de datos
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

' Función principal que ejecuta todas las pruebas del módulo
Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository"
    
    ' Ejecutar todas las pruebas de integración
    Call suiteResult.AddResult(TestGetUserAuthDataGeneric())
    
    Set TIAuthRepositoryRunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & LANZADERA_TEMPLATE_PATH, modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTestAuthRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testPath As String
    testPath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    If fs.FileExists(testPath) Then
        fs.DeleteFile testPath
    End If
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestGetUserAuthDataGeneric() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserAuthData debe funcionar sin errores y devolver objeto válido"
    
    Dim repository As IAuthRepository
    Dim errorHandler As IErrorHandlerService
    Dim testConfig As IConfig
    Dim fs As IFileSystem

    On Error GoTo TestError
    
    Setup
    
    Set testConfig = modConfigFactory.CreateConfigService()
    testConfig.SetSetting "LANZADERA_DB_PATH", modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    testConfig.SetSetting "LANZADERA_DB_PASSWORD", "dpddpd"
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    Set repository = modRepositoryFactory.CreateAuthRepository()
    
    Dim authData As EAuthData
    Set authData = repository.GetUserAuthData("test@example.com")
    
    modAssert.AssertNotNull authData, "GetUserAuthData debe devolver un objeto no nulo"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Teardown
    Set TestGetUserAuthDataGeneric = testResult
End Function

