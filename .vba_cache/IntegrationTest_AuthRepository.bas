Attribute VB_Name = "IntegrationTest_AuthRepository"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' MÓDULO DE PRUEBAS DE INTEGRACIÓN PARA CAuthRepository
' ============================================================================
' Este módulo contiene pruebas de integración para CAuthRepository
' que validan la conexión y consultas contra la base de datos real de la Lanzadera.

' Constantes para el autoaprovisionamiento de bases de datos
Private Const LANZADERA_TEMPLATE_PATH As String = "back\test_db\templates\Lanzadera_test_template.accdb"
Private Const LANZADERA_ACTIVE_PATH As String = "back\test_db\active\Lanzadera_integration_test.accdb"

' Función principal que ejecuta todas las pruebas del módulo
Public Function IntegrationTest_AuthRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_AuthRepository"
    
    ' Ejecutar todas las pruebas de integración
    suiteResult.AddTestResult Test_GetUserAuthData_Admin()
    suiteResult.AddTestResult Test_GetUserAuthData_Calidad()
    suiteResult.AddTestResult Test_GetUserAuthData_Tecnico()
    suiteResult.AddTestResult Test_GetUserAuthData_UserNotExists()
    
    Set IntegrationTest_AuthRepository_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

' Prepara el entorno de prueba copiando la plantilla de BD a BD de prueba
Private Sub Setup()
    On Error GoTo ErrorHandler
    
    ' Aprovisionar la base de datos de prueba usando el sistema estándar
    Dim fullTemplatePath As String
    Dim fullTestPath As String
    
    fullTemplatePath = modTestUtils.GetProjectPath() & LANZADERA_TEMPLATE_PATH
    fullTestPath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    
    modTestUtils.PrepareTestDatabase fullTemplatePath, fullTestPath
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Setup (" & Err.Number & "): " & Err.Description
    Err.Raise Err.Number, "IntegrationTest_AuthRepository.Setup", Err.Description
End Sub

' Limpia el entorno de prueba
Private Sub Teardown()
    On Error GoTo ErrorHandler
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testPath As String
    testPath = modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH
    
    ' Eliminar BD de prueba
    If fs.FileExists(testPath) Then
        fs.DeleteFile testPath
    End If
    
    Set fs = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Teardown: " & Err.Number & " - " & Err.Description
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

' Prueba que GetUserAuthData devuelve datos correctos para usuario Admin
Private Function Test_GetUserAuthData_Admin() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserAuthData_Admin"
    
    On Error GoTo ErrorHandler
    
    ' Setup
    Call Setup
    
    ' Arrange
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    
    ' Crear configuración específica para BD de prueba usando patrón de variable temporal
    Dim tempConfig As New CConfig
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH, "LANZADERA_DB_PATH"
    settings.Add "dpddpd", "LANZADERA_DB_PASSWORD"
    tempConfig.LoadFromCollection settings
    Set config = tempConfig
    
    Dim repository As New CAuthRepository
    repository.Initialize config, errorHandler
    
    ' Act
    Dim authData As T_AuthData
    Set authData = repository.GetUserAuthData("admin@test.com")
    
    ' Assert
    modAssert.AssertTrue authData.UserExists, "Usuario admin debe existir"
    modAssert.AssertTrue authData.IsGlobalAdmin, "Usuario admin debe tener rol de administrador global"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Set fileSystem = Nothing
    Set Test_GetUserAuthData_Admin = testResult
End Function

' Prueba que GetUserAuthData devuelve datos correctos para usuario de Calidad
Private Function Test_GetUserAuthData_Calidad() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserAuthData_Calidad"
    
    On Error GoTo ErrorHandler
    
    ' Setup
    Call Setup
    
    ' Arrange
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    
    ' Crear configuración específica para BD de prueba usando patrón de variable temporal
    Dim tempConfig As New CConfig
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH, "LANZADERA_DB_PATH"
    settings.Add "dpddpd", "LANZADERA_DB_PASSWORD"
    tempConfig.LoadFromCollection settings
    Set config = tempConfig
    
    Dim repository As New CAuthRepository
    repository.Initialize config, errorHandler
    
    ' Act
    Dim authData As T_AuthData
    Set authData = repository.GetUserAuthData("calidad@test.com")
    
    ' Assert
    modAssert.AssertTrue authData.UserExists, "Usuario calidad debe existir"
    modAssert.AssertTrue authData.IsCalidad, "Usuario calidad debe tener rol de calidad"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario calidad no debe ser administrador global"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Set fileSystem = Nothing
    Set Test_GetUserAuthData_Calidad = testResult
End Function

' Prueba que GetUserAuthData devuelve datos correctos para usuario Técnico
Private Function Test_GetUserAuthData_Tecnico() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserAuthData_Tecnico"
    
    On Error GoTo ErrorHandler
    
    ' Setup
    Call Setup
    
    ' Arrange
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    
    ' Crear configuración específica para BD de prueba usando patrón de variable temporal
    Dim tempConfig As New CConfig
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH, "LANZADERA_DB_PATH"
    settings.Add "dpddpd", "LANZADERA_DB_PASSWORD"
    tempConfig.LoadFromCollection settings
    Set config = tempConfig
    
    Dim repository As New CAuthRepository
    repository.Initialize config, errorHandler
    
    ' Act
    Dim authData As T_AuthData
    Set authData = repository.GetUserAuthData("tecnico@test.com")
    
    ' Assert
    modAssert.AssertTrue authData.UserExists, "Usuario técnico debe existir"
    modAssert.AssertTrue authData.IsTecnico, "Usuario técnico debe tener rol técnico"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario técnico no debe ser administrador global"
    modAssert.AssertFalse authData.IsCalidad, "Usuario técnico no debe tener rol de calidad"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Set fileSystem = Nothing
    Set Test_GetUserAuthData_Tecnico = testResult
End Function

' Prueba que GetUserAuthData maneja correctamente usuarios inexistentes
Private Function Test_GetUserAuthData_UserNotExists() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetUserAuthData_UserNotExists"
    
    On Error GoTo ErrorHandler
    
    ' Setup
    Call Setup
    
    ' Arrange
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    
    Dim config As IConfig
    Set config = modConfig.CreateConfigService()
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    
    ' Crear configuración específica para BD de prueba usando patrón de variable temporal
    Dim tempConfig As New CConfig
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & LANZADERA_ACTIVE_PATH, "LANZADERA_DB_PATH"
    settings.Add "dpddpd", "LANZADERA_DB_PASSWORD"
    tempConfig.LoadFromCollection settings
    Set config = tempConfig
    
    Dim repository As New CAuthRepository
    repository.Initialize config, errorHandler
    
    ' Act
    Dim authData As T_AuthData
    Set authData = repository.GetUserAuthData("inexistente@test.com")
    
    ' Assert
    modAssert.AssertFalse authData.UserExists, "Usuario inexistente no debe existir"
    modAssert.AssertFalse authData.IsGlobalAdmin, "Usuario inexistente no debe ser administrador global"
    modAssert.AssertFalse authData.IsCalidad, "Usuario inexistente no debe tener rol de calidad"
    modAssert.AssertFalse authData.IsTecnico, "Usuario inexistente no debe tener rol técnico"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Call Teardown
    Set repository = Nothing
    Set config = Nothing
    Set errorHandler = Nothing
    Set fileSystem = Nothing
    Set Test_GetUserAuthData_UserNotExists = testResult
End Function

#End If