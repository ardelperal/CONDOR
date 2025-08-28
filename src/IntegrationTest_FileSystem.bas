Attribute VB_Name = "IntegrationTest_FileSystem"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: IntegrationTest_FileSystem
' DESCRIPCIÓN: Pruebas de integración para el módulo FileSystem
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Suite de pruebas de integración
' =====================================================

' Suite de pruebas de integración para validar el comportamiento
' del módulo FileSystem con el sistema de archivos real

Private Const TEST_ROOT_PATH As String = "back\test_env\fs_test"



' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function IntegrationTest_FileSystem_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_FileSystem"
    
    ' Configurar entorno
    Setup
    
    ' Ejecutar tests individuales
    suiteResult.AddTestResult Test_FileExists_Integration()
    suiteResult.AddTestResult Test_FolderExists_Integration()
    suiteResult.AddTestResult Test_CopyFile_Integration()
    suiteResult.AddTestResult Test_DeleteFile_Integration()
    suiteResult.AddTestResult Test_CreateFolder_Integration()
    suiteResult.AddTestResult Test_DeleteFolder_Integration()
    
    ' Limpiar entorno
    Teardown
    
    Set IntegrationTest_FileSystem_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Debug.Print "[SETUP] Preparando entorno de prueba FileSystem..."
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Eliminar directorio de prueba si existe
    If fso.FolderExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH) Then
        fso.DeleteFolder modTestUtils.GetProjectPath() & TEST_ROOT_PATH, True
    End If
    
    ' Crear directorio de prueba limpio
    fso.CreateFolder modTestUtils.GetProjectPath() & TEST_ROOT_PATH
    
    Debug.Print "[SETUP] Entorno preparado en: " & modTestUtils.GetProjectPath() & TEST_ROOT_PATH
    Set fso = Nothing
End Sub

Private Sub Teardown()
    Debug.Print "[TEARDOWN] Limpiando entorno de prueba FileSystem..."
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Eliminar directorio de prueba y todo su contenido
    Dim testRootPath As String
    testRootPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH
    
    If fs.FolderExists(testRootPath) Then
        fs.DeleteFolder testRootPath
    End If
    
    Debug.Print "[TEARDOWN] Entorno limpiado"
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN INDIVIDUALES
' ============================================================================

Private Function Test_FileExists_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_FileExists_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_FileExists_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testFilePath As String
    testFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\test_file.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear archivo de prueba
    Dim textFile As Object
    Set textFile = fso.CreateTextFile(testFilePath, True)
    textFile.WriteLine "Contenido de prueba"
    textFile.Close
    
    ' Act & Assert
    modAssert.AssertTrue fs.FileExists(testFilePath), "El archivo debería existir"
    modAssert.AssertFalse fs.FileExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\archivo_inexistente.txt"), "El archivo inexistente no debería existir"
    
    testResult.Pass
    Debug.Print "[TEST] Test_FileExists_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_FileExists_Integration = testResult
End Function

Private Function Test_FolderExists_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_FolderExists_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_FolderExists_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\test_folder"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta de prueba
    fso.CreateFolder testFolderPath
    
    ' Act & Assert
    modAssert.AssertTrue fs.FolderExists(testFolderPath), "La carpeta debería existir"
    modAssert.AssertFalse fs.FolderExists(modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\carpeta_inexistente"), "La carpeta inexistente no debería existir"
    
    testResult.Pass
    Debug.Print "[TEST] Test_FolderExists_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_FolderExists_Integration = testResult
End Function

Private Function Test_CopyFile_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_CopyFile_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_CopyFile_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim sourceFilePath As String
    Dim destinationFilePath As String
    sourceFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\source_file.txt"
    destinationFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\destination_file.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear archivo origen
    Dim textFile As Object
    Set textFile = fso.CreateTextFile(sourceFilePath, True)
    textFile.WriteLine "Contenido para copiar"
    textFile.Close
    
    ' Act
    fs.CopyFile sourceFilePath, destinationFilePath
    
    ' Assert
    modAssert.AssertTrue fso.FileExists(destinationFilePath), "El archivo de destino debería existir después de la copia"
    
    testResult.Pass
    Debug.Print "[TEST] Test_CopyFile_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_CopyFile_Integration = testResult
End Function

Private Function Test_DeleteFile_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_DeleteFile_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_DeleteFile_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testFilePath As String
    testFilePath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\file_to_delete.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear archivo para eliminar
    Dim textFile As Object
    Set textFile = fso.CreateTextFile(testFilePath, True)
    textFile.WriteLine "Archivo para eliminar"
    textFile.Close
    
    ' Verificar que existe antes de eliminar
    modAssert.AssertTrue fso.FileExists(testFilePath), "El archivo debería existir antes de eliminarlo"
    
    ' Act
    fs.DeleteFile testFilePath
    
    ' Assert
    modAssert.AssertFalse fso.FileExists(testFilePath), "El archivo no debería existir después de eliminarlo"
    
    testResult.Pass
    Debug.Print "[TEST] Test_DeleteFile_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_DeleteFile_Integration = testResult
End Function

Private Function Test_CreateFolder_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_CreateFolder_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_CreateFolder_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\new_folder"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar que no existe antes de crear
    modAssert.AssertFalse fso.FolderExists(testFolderPath), "La carpeta no debería existir antes de crearla"
    
    ' Act
    fs.CreateFolder testFolderPath
    
    ' Assert
    modAssert.AssertTrue fso.FolderExists(testFolderPath), "La carpeta debería existir después de crearla"
    
    testResult.Pass
    Debug.Print "[TEST] Test_CreateFolder_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_CreateFolder_Integration = testResult
End Function

Private Function Test_DeleteFolder_Integration() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_DeleteFolder_Integration"
    
    On Error GoTo ErrorHandler
    
    Debug.Print "[TEST] Test_DeleteFolder_Integration"
    
    ' Arrange
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    Dim testFolderPath As String
    testFolderPath = modTestUtils.GetProjectPath() & TEST_ROOT_PATH & "\folder_to_delete"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta para eliminar
    fso.CreateFolder testFolderPath
    
    ' Verificar que existe antes de eliminar
    modAssert.AssertTrue fso.FolderExists(testFolderPath), "La carpeta debería existir antes de eliminarla"
    
    ' Act
    fs.DeleteFolder testFolderPath
    
    ' Assert
    modAssert.AssertFalse fso.FolderExists(testFolderPath), "La carpeta no debería existir después de eliminarla"
    
    testResult.Pass
    Debug.Print "[TEST] Test_DeleteFolder_Integration - PASÓ"
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error: " & Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
    Set Test_DeleteFolder_Integration = testResult
End Function