Attribute VB_Name = "TIFileSystem"
Option Compare Database
Option Explicit

Private Const TEST_DIR As String = "condor_test_fs"
Private Const TEST_FILE As String = "test_file.txt"

Private Function GetTestPath() As String
    GetTestPath = modTestUtils.GetProjectPath & TEST_DIR
End Function

Public Function TIFileSystemRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIFileSystem"
    
    suiteResult.AddResult TestCreateAndFolderExists()
    suiteResult.AddResult TestCreateAndDeleteFile()
    
    Set TIFileSystemRunAll = suiteResult
End Function

Private Sub Setup()
    On Error Resume Next
    Teardown
    On Error GoTo 0
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.CreateFolder GetTestPath
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFolderRecursive GetTestPath
End Sub

Private Function TestCreateAndFolderExists() As CTestResult
    Set TestCreateAndFolderExists = New CTestResult
    TestCreateAndFolderExists.Initialize "CreateFolder y FolderExists deben funcionar"
    
    Dim fs As IFileSystem
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Assert
    modAssert.AssertTrue fs.FolderExists(GetTestPath), "La carpeta de prueba debería existir después del Setup."
    
    TestCreateAndFolderExists.Pass
    GoTo Cleanup
    
TestFail:
    TestCreateAndFolderExists.Fail Err.Description
    
Cleanup:
    Call Teardown
End Function

Private Function TestCreateAndDeleteFile() As CTestResult
    Set TestCreateAndDeleteFile = New CTestResult
    TestCreateAndDeleteFile.Initialize "CopyFile, FileExists y DeleteFile deben funcionar"
    
    Dim fs As IFileSystem
    Dim testFilePath As String
    On Error GoTo TestFail
    
    ' Arrange
    Call Setup
    Set fs = modFileSystemFactory.CreateFileSystem()
    testFilePath = GetTestPath & "\" & TEST_FILE
    
    ' Crear un fichero dummy para copiar
    Dim tempSourcePath As String
    tempSourcePath = GetTestPath & "\source.txt"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile(tempSourcePath, True).Write "test"
    
    ' Act
    fs.CopyFile tempSourcePath, testFilePath
    
    ' Assert
    modAssert.AssertTrue fs.FileExists(testFilePath), "El archivo copiado debería existir."
    fs.DeleteFile testFilePath
    modAssert.AssertFalse fs.FileExists(testFilePath), "El archivo eliminado no debería existir."
    
    TestCreateAndDeleteFile.Pass
    GoTo Cleanup
    
TestFail:
    TestCreateAndDeleteFile.Fail Err.Description
    
Cleanup:
    Call Teardown
End Function


