Attribute VB_Name = "TIWordManager"
Option Compare Database
Option Explicit


' ============================================================================
' Módulo: TIWordManager
' Descripción: Pruebas de integración para CWordManager, siguiendo el
'              patrón de auto-aprovisionamiento (Lección 36).
' ============================================================================

Private Const TEST_FOLDER_NAME As String = "condor_word_tests"
Private Const TEMPLATE_DOC_NAME As String = "template_test.docx"
Private Const MODIFIED_DOC_NAME As String = "modified_test.docx"

Private Function GetTestFolderPath() As String
    GetTestFolderPath = Environ("TEMP") & "\" & TEST_FOLDER_NAME & "\"
End Function

Public Function TIWordManager_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestWordManager"
    
    suiteResult.AddResult Test_CicloCompleto_Success()
    suiteResult.AddResult Test_AbrirFicheroInexistente_DevuelveFalse()
    
    Set TIWordManager_RunAll = suiteResult
End Function

Private Function Test_CicloCompleto_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Ciclo completo (Abrir, Reemplazar, Guardar, Leer) debe tener éxito"
    
    Dim fs As IFileSystem
    Dim wordManager As IWordManager
    On Error GoTo TestFail
    
    Call SetupTestEnvironment
    Set fs = modFileSystemFactory.CreateFileSystem
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim templatePath As String: templatePath = GetTestFolderPath & TEMPLATE_DOC_NAME
    
    modAssert.AssertTrue wordManager.AbrirDocumento(templatePath), "Debería abrir el documento"
    modAssert.AssertTrue wordManager.ReemplazarTexto("[NOMBRE]", "CONDOR"), "Debería reemplazar el texto"
    
    Dim modifiedPath As String: modifiedPath = GetTestFolderPath & MODIFIED_DOC_NAME
    modAssert.AssertTrue wordManager.GuardarDocumento(modifiedPath), "Debería guardar el documento"
    wordManager.CerrarDocumento
    
    modAssert.AssertTrue fs.FileExists(modifiedPath), "El archivo modificado debe existir"
    Dim finalContent As String: finalContent = wordManager.LeerContenidoDocumento(modifiedPath)
    
    modAssert.AssertTrue InStr(1, finalContent, "CONDOR", vbTextCompare) > 0, "El contenido debe incluir 'CONDOR'"
    modAssert.AssertTrue InStr(1, finalContent, "[NOMBRE]", vbTextCompare) = 0, "El contenido no debe incluir '[NOMBRE]'"
    
    testResult.Pass
Finally:
    Call TeardownTestEnvironment
    Set fs = Nothing
    Set wordManager = Nothing
    Set Test_CicloCompleto_Success = testResult
    Exit Function
TestFail:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Finally
End Function

Private Function Test_AbrirFicheroInexistente_DevuelveFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "AbrirDocumento con un fichero inexistente debe devolver False"
    
    Dim wordManager As IWordManager
    On Error GoTo TestFail
    
    Call SetupTestEnvironment
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim result As Boolean: result = wordManager.AbrirDocumento(GetTestFolderPath & "non_existent_file.docx")
    
    modAssert.AssertFalse result, "Debería devolver False"
    testResult.Pass
Finally:
    Call TeardownTestEnvironment
    Set wordManager = Nothing
    Set Test_AbrirFicheroInexistente_DevuelveFalse = testResult
    Exit Function
TestFail:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Finally
End Function

Private Sub SetupTestEnvironment()
    On Error Resume Next
    Call TeardownTestEnvironment
    On Error GoTo 0
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem
    fs.CreateFolder GetTestFolderPath
    
    CreateDummyWordDocument GetTestFolderPath & TEMPLATE_DOC_NAME, "Hola [NOMBRE], bienvenido."
End Sub

Private Sub TeardownTestEnvironment()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem
    If fs.FolderExists(GetTestFolderPath) Then fs.DeleteFolder GetTestFolderPath
End Sub

Private Sub CreateDummyWordDocument(ByVal filePath As String, ByVal content As String)
    Dim wordApp As Object, wordDoc As Object
    On Error GoTo Cleanup
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add
    wordDoc.content.Text = content
    wordDoc.SaveAs2 filePath
    
Cleanup:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub


