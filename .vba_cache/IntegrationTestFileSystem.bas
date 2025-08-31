Attribute VB_Name = "IntegrationTestFileSystem"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: IntegrationTestFileSystem
' DESCRIPCIÓN: Pruebas de integración para el módulo FileSystem
' =====================================================

Private Const TEST_ROOT_PATH As String = "back\test_env\fs_test"

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function IntegrationTestFileSystemRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestFileSystem"
    
    suiteResult.AddTestResult TestFileExistsIntegration()
    suiteResult.AddTestResult TestFolderExistsIntegration()
    suiteResult.AddTestResult TestCopyFileIntegration()
    suiteResult.AddTestResult TestDeleteFileIntegration()
    suiteResult.AddTestResult TestCreateFolderIntegration()
    suiteResult.AddTestResult TestDeleteFolderIntegration()
    
    Set IntegrationTestFileSystemRunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH) Then
        fso.DeleteFolder modTestUtils.GetProjectPath() & TEST_ROOT_PATH, True
    End If
    fso.CreateFolder modTestUtils.GetProjectPath() & TEST_ROOT_PATH
    Set fso = Nothing
End Sub

Private Sub Teardown()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testRootPath As String
    testRootPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH
    If fs.FolderExists(testRootPath) Then
        fs.DeleteFolder testRootPath
    End If
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN INDIVIDUALES
' ============================================================================

Private Function TestFileExistsIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "FileExists debe detectar correctamente archivos existentes e inexistentes"
    On Error GoTo ErrorHandler
    
    Setup
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFilePath As String
    testFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\test_file.txt"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile(testFilePath, True).Close
    
    modAssert.AssertTrue fs.FileExists(testFilePath), "El archivo debería existir"
    modAssert.AssertFalse fs.FileExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\archivo_inexistente.txt"), "El archivo inexistente no debería existir"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestFileExistsIntegration = testResult
End Function

Private Function TestFolderExistsIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "FolderExists debe detectar correctamente carpetas existentes e inexistentes"
    On Error GoTo ErrorHandler

    Setup

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\test_folder"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder testFolderPath
    
    modAssert.AssertTrue fs.FolderExists(testFolderPath), "La carpeta debería existir"
    modAssert.AssertFalse fs.FolderExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\carpeta_inexistente"), "La carpeta inexistente no debería existir"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestFolderExistsIntegration = testResult
End Function

Private Function TestCopyFileIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CopyFile debe copiar un archivo correctamente"
    On Error GoTo ErrorHandler

    Setup

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim sourceFilePath As String
    Dim destinationFilePath As String
    sourceFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\source_file.txt"
    destinationFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\destination_file.txt"
    CreateObject("Scripting.FileSystemObject").CreateTextFile(sourceFilePath, True).Close
    
    fs.CopyFile sourceFilePath, destinationFilePath
    
    modAssert.AssertTrue fs.FileExists(destinationFilePath), "El archivo de destino debería existir después de la copia"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestCopyFileIntegration = testResult
End Function

Private Function TestDeleteFileIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "DeleteFile debe eliminar un archivo correctamente"
    On Error GoTo ErrorHandler

    Setup

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFilePath As String
    testFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\file_to_delete.txt"
    CreateObject("Scripting.FileSystemObject").CreateTextFile(testFilePath, True).Close
    
    modAssert.AssertTrue fs.FileExists(testFilePath), "El archivo debería existir antes de eliminarlo"
    
    fs.DeleteFile testFilePath
    
    modAssert.AssertFalse fs.FileExists(testFilePath), "El archivo no debería existir después de eliminarlo"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestDeleteFileIntegration = testResult
End Function

Private Function TestCreateFolderIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CreateFolder debe crear una carpeta correctamente"
    On Error GoTo ErrorHandler

    Setup

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\new_folder"
    
    modAssert.AssertFalse fs.FolderExists(testFolderPath), "La carpeta no debería existir antes de crearla"
    
    fs.CreateFolder testFolderPath
    
    modAssert.AssertTrue fs.FolderExists(testFolderPath), "La carpeta debería existir después de crearla"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestCreateFolderIntegration = testResult
End Function

Private Function TestDeleteFolderIntegration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "DeleteFolder debe eliminar una carpeta correctamente"
    On Error GoTo ErrorHandler

    Setup

    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\folder_to_delete"
    CreateObject("Scripting.FileSystemObject").CreateFolder testFolderPath
    
    modAssert.AssertTrue fs.FolderExists(testFolderPath), "La carpeta debería existir antes de eliminarla"
    
    fs.DeleteFolder testFolderPath
    
    modAssert.AssertFalse fs.FolderExists(testFolderPath), "La carpeta no debería existir después de eliminarla"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
Cleanup:
    Teardown
    Set TestDeleteFolderIntegration = testResult
End Function
