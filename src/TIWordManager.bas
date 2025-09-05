Attribute VB_Name = "TIWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' REQUISITO DE COMPILACIÓN CRÍTICO:
' "Microsoft Word XX.X Object Library" debe estar referenciada.
' ============================================================================

Private Const TEST_FOLDER_PATH As String = "back\test_db\active\word_manager_tests\"
Private Const TEMPLATE_DOC_NAME As String = "template_test.docx"
Private Const MODIFIED_DOC_NAME As String = "modified_test.docx"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWordManager (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult Test_CicloCompleto_Success()
    suiteResult.AddResult Test_AbrirFicheroInexistente_DevuelveFalse()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIWordManagerRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFolder As String: testFolder = modTestUtils.GetProjectPath() & TEST_FOLDER_PATH
    
    ' Crear directorio de prueba limpio
    If fs.FolderExists(testFolder) Then
        fs.DeleteFolderRecursive testFolder
    End If
    fs.CreateFolder testFolder
    
    ' Crear el fichero de plantilla de Word mediante programación
    Call CreateTestTemplate(testFolder & TEMPLATE_DOC_NAME)
End Sub

Private Sub SuiteTeardown()
    On Error Resume Next ' Blindaje
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFolderRecursive modTestUtils.GetProjectPath() & TEST_FOLDER_PATH
    Set fs = Nothing
End Sub

Private Sub CreateTestTemplate(ByVal templatePath As String)
    Dim wordApp As Object, doc As Object
    On Error GoTo ErrorHandler
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Set doc = wordApp.Documents.Add
    doc.Content.Text = "Hola [NOMBRE], este es un documento de prueba."
    doc.SaveAs2 templatePath
    
ErrorHandler:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=0 ' wdDoNotSaveChanges
    If Not wordApp Is Nothing Then wordApp.Quit
    Set doc = Nothing
    Set wordApp = Nothing
End Sub

' ============================================================================
' TESTS INDIVIDUALES (SE AÑADIRÁN EN LOS SIGUIENTES PROMPTS)
' ============================================================================

Private Function Test_CicloCompleto_Success() As CTestResult
    Set Test_CicloCompleto_Success = New CTestResult
    Test_CicloCompleto_Success.Initialize "Ciclo completo (Abrir, Reemplazar, Guardar, Leer) debe tener éxito"
    
    Dim wordManager As IWordManager
    Dim fs As IFileSystem
    On Error GoTo TestFail
    
    ' ARRANGE - Crear configuración local para tests de integración
    Dim localConfig As IConfig
    Set localConfig = New CConfig
    localConfig.SetValue "DatabasePath", TEST_DB_PATH
    localConfig.SetValue "DatabasePassword", TEST_DB_PASSWORD
    localConfig.SetValue "OutputPath", TEST_OUTPUT_PATH
    
    Set wordManager = modWordManagerFactory.CreateWordManager(localConfig)
    Set fs = modFileSystemFactory.CreateFileSystem(localConfig)
    Dim templatePath As String: templatePath = modTestUtils.GetProjectPath() & TEST_FOLDER_PATH & TEMPLATE_DOC_NAME
    Dim modifiedPath As String: modifiedPath = modTestUtils.GetProjectPath() & TEST_FOLDER_PATH & MODIFIED_DOC_NAME
    
    ' ACT
    wordManager.AbrirDocumento templatePath
    wordManager.ReemplazarTexto "[NOMBRE]", "CONDOR"
    wordManager.GuardarDocumento modifiedPath
    Dim contenido As String: contenido = wordManager.LeerContenidoDocumento(modifiedPath)
    
    ' ASSERT
    modAssert.AssertTrue InStr(contenido, "CONDOR") > 0, "El contenido debería contener 'CONDOR'"
    modAssert.AssertTrue fs.FileExists(modifiedPath), "El fichero modificado debe existir."
    
    Test_CicloCompleto_Success.Pass
    GoTo Cleanup

TestFail:
    Test_CicloCompleto_Success.Fail "Error: " & Err.Description
    
Cleanup:
    If Not wordManager Is Nothing Then wordManager.Dispose
    Set fs = Nothing
    Set wordManager = Nothing
End Function

Private Function Test_AbrirFicheroInexistente_DevuelveFalse() As CTestResult
    Set Test_AbrirFicheroInexistente_DevuelveFalse = New CTestResult
    Test_AbrirFicheroInexistente_DevuelveFalse.Initialize "Abrir un fichero inexistente debe devolver False y no lanzar error"
    
    Dim wordManager As IWordManager
    On Error GoTo TestFail ' Si hay algún error, el test falla
    
    ' Arrange - Crear configuración local para tests de integración
    Dim localConfig As IConfig
    Set localConfig = New CConfig
    localConfig.SetValue "DatabasePath", TEST_DB_PATH
    localConfig.SetValue "DatabasePassword", TEST_DB_PASSWORD
    localConfig.SetValue "OutputPath", TEST_OUTPUT_PATH
    
    Set wordManager = modWordManagerFactory.CreateWordManager(localConfig)
    Dim inexistentePath As String: inexistentePath = modTestUtils.GetProjectPath() & TEST_FOLDER_PATH & "no_existe.docx"
    
    ' Act
    Dim result As Boolean
    result = wordManager.AbrirDocumento(inexistentePath)
    
    ' Assert
    modAssert.AssertFalse result, "AbrirDocumento debería haber devuelto False."
    
    Test_AbrirFicheroInexistente_DevuelveFalse.Pass
    GoTo Cleanup

TestFail:
    Test_AbrirFicheroInexistente_DevuelveFalse.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    If Not wordManager Is Nothing Then wordManager.Dispose
    Set wordManager = Nothing
End Function