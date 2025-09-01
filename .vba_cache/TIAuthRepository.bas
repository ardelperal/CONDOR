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
    Call suiteResult.AddResult(TestGetUserAuthDataAdmin())
    Call suiteResult.AddResult(TestGetUserAuthDataCalidad())
    Call suiteResult.AddResult(TestGetUserAuthDataTecnico())
    Call suiteResult.AddResult(TestGetUserAuthDataUserNotExists())
    
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

Private Function TestGetUserAuthDataAdmin() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserAuthData debe devolver datos correctos para un usuario Admin"
    
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
    Set authData = repository.GetUserAuthData("admin@test.com")
    
    modAssert.AssertTrue authData.UserExists, "Usuario admin debe existir"
    modAssert.AssertTrue authData.IsGlobalAdmin, "Usuario admin debe tener rol de administrador global"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Teardown
    Set TestGetUserAuthDataAdmin = testResult
End Function

Private Function TestGetUserAuthDataCalidad() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserAuthData debe devolver datos correctos para un usuario de Calidad"
    
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
     Set authData = repository.GetUserAuthData("calidad@test.com")
    
    modAssert.AssertTrue authData.UserExists, "Usuario calidad debe existir"
    modAssert.AssertTrue authData.IsCalidad, "Usuario calidad debe tener rol de calidad"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario calidad no debe ser administrador global"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Teardown
    Set TestGetUserAuthDataCalidad = testResult
End Function

Private Function TestGetUserAuthDataTecnico() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserAuthData debe devolver datos correctos para un usuario Técnico"
    
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
     Set authData = repository.GetUserAuthData("tecnico@test.com")
    
    modAssert.AssertTrue authData.UserExists, "Usuario técnico debe existir"
    modAssert.AssertTrue authData.IsTecnico, "Usuario técnico debe tener rol técnico"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario técnico no debe ser administrador global"
    modAssert.AssertFalse authData.IsCalidad, "Usuario técnico no debe tener rol de calidad"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Teardown
    Set TestGetUserAuthDataTecnico = testResult
End Function

Private Function TestGetUserAuthDataUserNotExists() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetUserAuthData debe manejar correctamente usuarios inexistentes"
    
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
     Set authData = repository.GetUserAuthData("inexistente@test.com")
    
    modAssert.AssertFalse authData.UserExists, "Usuario inexistente no debe existir"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario inexistente no debe ser administrador global"
    modAssert.AssertFalse authData.IsCalidad, "Usuario inexistente no debe tener rol de calidad"
    modAssert.AssertFalse authData.IsTecnico, "Usuario inexistente no debe tener rol técnico"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Teardown
    Set TestGetUserAuthDataUserNotExists = testResult
End Function