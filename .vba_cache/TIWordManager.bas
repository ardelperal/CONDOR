Attribute VB_Name = "TIWordManager"
Option Compare Database
Option Explicit
' ============================================================================
' ¡¡¡ REQUISITO DE COMPILACIÓN CRÍTICO !!!
' Este módulo interactúa con Microsoft Word y requiere una referencia externa
' para compilar y funcionar correctamente.
'
' ACCIÓN REQUERIDA: En el editor de VBA, vaya a:
' Herramientas -> Referencias...
' Y asegúrese de que la siguiente casilla esté marcada:
' "Microsoft Word XX.X Object Library" (la versión puede variar)
' ============================================================================


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

Public Function TIWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWordManager"
    
    suiteResult.AddResult Test_CicloCompleto_Success()
    suiteResult.AddResult Test_AbrirFicheroInexistente_DevuelveFalse()
    
    Set TIWordManagerRunAll = suiteResult
End Function

Private Function Test_CicloCompleto_Success() As CTestResult
    Set Test_CicloCompleto_Success = New CTestResult
    Test_CicloCompleto_Success.Initialize "Ciclo completo (Abrir, Reemplazar, Guardar, Leer) debe tener éxito"
    
    Dim fs As IFileSystem
    Dim wordManager As IWordManager
    
    On Error GoTo TestFail
    
    ' ARRANGE: Preparar el entorno y los objetos
    Call SetupTestEnvironment
    Set fs = modFileSystemFactory.CreateFileSystem
    Set wordManager = modWordManagerFactory.CreateWordManager()
    Dim templatePath As String: templatePath = GetTestFolderPath & TEMPLATE_DOC_NAME
    Dim modifiedPath As String: modifiedPath = GetTestFolderPath & MODIFIED_DOC_NAME
    
    ' ACT: Ejecutar la secuencia de operaciones
    wordManager.AbrirDocumento templatePath
    wordManager.ReemplazarTexto "[NOMBRE]", "CONDOR"
    wordManager.GuardarDocumento modifiedPath
    wordManager.CerrarDocumento
    Dim contenido As String: contenido = wordManager.LeerContenidoDocumento(modifiedPath)
    
    ' ASSERT: Verificar los resultados
    modAssert.AssertTrue InStr(contenido, "CONDOR") > 0, "El contenido debería contener 'CONDOR'"
    
    Test_CicloCompleto_Success.Pass
    GoTo Cleanup

TestFail:
    Test_CicloCompleto_Success.Fail "Error: " & Err.Description
    
Cleanup:
    ' Bloque de limpieza único y garantizado
    Call CleanupTestEnvironment
    Set fs = Nothing
    Set wordManager = Nothing
End Function

Private Function Test_AbrirFicheroInexistente_DevuelveFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Abrir fichero inexistente debe devolver False"
    
    Dim wordManager As IWordManager
    On Error GoTo TestFail
    
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    Dim inexistentePath As String: inexistentePath = "C:\fichero_que_no_existe.docx"
    modAssert.AssertFalse wordManager.AbrirDocumento(inexistentePath), "No debería abrir un fichero inexistente"
    
    testResult.Pass
    GoTo TestEnd
    
TestFail:
    testResult.Fail "Error: " & Err.Description
    
TestEnd:
    Set Test_AbrirFicheroInexistente_DevuelveFalse = testResult
End Function

Private Sub SetupTestEnvironment()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem
    
    Dim testFolder As String: testFolder = GetTestFolderPath
    If Not fs.FolderExists(testFolder) Then
        fs.CreateFolder testFolder
    End If
    
    Dim templatePath As String: templatePath = testFolder & TEMPLATE_DOC_NAME
    If Not fs.FileExists(templatePath) Then
        Call CreateTestTemplate(templatePath)
    End If
End Sub

Private Sub CleanupTestEnvironment()
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem
    
    Dim testFolder As String: testFolder = GetTestFolderPath
    If fs.FolderExists(testFolder) Then
        fs.DeleteFolderRecursive testFolder
    End If
End Sub

Private Sub CreateTestTemplate(templatePath As String)
    Dim wordApp As Object
    Dim doc As Object
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Set doc = wordApp.Documents.Add
    doc.Content.Text = "Hola [NOMBRE], este es un documento de prueba."
    doc.SaveAs2 templatePath
    doc.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing
End Sub