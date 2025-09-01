Attribute VB_Name = "TIWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: IntegrationTestWordManager
' Descripción: Pruebas de integración para CWordManager
' ============================================================================

' Variables eliminadas - ahora se declaran localmente en cada función

' ============================================================================
' CONFIGURACIÓN DE PRUEBAS
' ============================================================================

Public Function TIWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWordManager"
    
    suiteResult.AddTestResult IntegrationTestWordManagerCicloCompletoSuccess()
    suiteResult.AddTestResult IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse()
    
    Set TIWordManagerRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function IntegrationTestWordManagerCicloCompletoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Ciclo completo de WordManager (Abrir, Reemplazar, Guardar, Cerrar) debe tener éxito"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim archivoOriginal As String
    Dim archivoGuardado As String
    Dim contenidoFinal As String
    Dim tempFolder As String
    Dim fs As IFileSystem
    
    ' Setup local
    tempFolder = modTestUtils.GetProjectPath() & "back\test_env\word_tests\"
    Set fs = modFileSystemFactory.CreateFileSystem()
    If Not fs.FolderExists(tempFolder) Then
        fs.CreateFolder tempFolder
    End If
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    archivoOriginal = tempFolder & "documento_original.docx"
    CrearDocumentoPrueba archivoOriginal, "Hola [NOMBRE], este es un documento de prueba."
    
    archivoGuardado = tempFolder & "documento_modificado.docx"
    
    ' Act & Assert
    modAssert.AssertTrue wordManager.AbrirDocumento(archivoOriginal), "Debería abrir el documento correctamente"
    modAssert.AssertTrue wordManager.ReemplazarTexto("[NOMBRE]", "CONDOR"), "Debería reemplazar el texto correctamente"
    modAssert.AssertTrue wordManager.GuardarDocumento(archivoGuardado), "Debería guardar el documento correctamente"
    wordManager.CerrarDocumento
    
    contenidoFinal = wordManager.LeerContenidoDocumento(archivoGuardado)
    modAssert.AssertTrue InStr(contenidoFinal, "CONDOR") > 0, "El contenido debería incluir 'CONDOR'"
    modAssert.AssertTrue InStr(contenidoFinal, "[ NOMBRE]") = 0, "El contenido no debería incluir '[NOMBRE]'"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    wordManager.CerrarDocumento
    On Error GoTo 0
    
Cleanup:
    ' Cleanup local
    On Error Resume Next
    If fs.FileExists(archivoOriginal) Then fs.DeleteFile archivoOriginal
    If fs.FileExists(archivoGuardado) Then fs.DeleteFile archivoGuardado
    If fs.FolderExists(tempFolder) Then fs.DeleteFolder tempFolder
    Set fs = Nothing
    On Error GoTo 0
    Set IntegrationTestWordManagerCicloCompletoSuccess = testResult
End Function

Private Function IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "AbrirDocumento con un fichero inexistente debe devolver False"
    
    On Error GoTo TestError
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim rutaInvalida As String
    Dim resultado As Boolean
    Dim tempFolder As String
    Dim fs As IFileSystem
    
    ' Setup local
    tempFolder = modTestUtils.GetProjectPath() & "back\test_env\word_tests\"
    Set fs = modFileSystemFactory.CreateFileSystem()
    If Not fs.FolderExists(tempFolder) Then
        fs.CreateFolder tempFolder
    End If
    
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    rutaInvalida = tempFolder & "archivo_que_no_existe.docx"
    
    ' Act
    resultado = wordManager.AbrirDocumento(rutaInvalida)
    
    ' Assert
    modAssert.AssertFalse resultado, "Debería devolver False al intentar abrir un archivo inexistente"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    
Cleanup:
    ' Cleanup local
    On Error Resume Next
    If fs.FolderExists(tempFolder) Then fs.DeleteFolder tempFolder
    Set fs = Nothing
    On Error GoTo 0
    Set IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse = testResult
End Function

' ============================================================================
' MÉTODOS DE SETUP Y TEARDOWN CENTRALIZADOS
' ============================================================================



' ============================================================================
' MÉTODOS AUXILIARES
' ============================================================================

Private Sub CrearDocumentoPrueba(ByVal rutaArchivo As String, ByVal contenido As String)
    On Error GoTo TestError
    Dim wordApp As Object
    Dim wordDoc As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    Set wordDoc = wordApp.Documents.Add
    wordDoc.content.Text = contenido
    wordDoc.SaveAs2 rutaArchivo, 16 ' wdFormatXMLDocument
    wordDoc.Close False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Sub
TestError:
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close False
        Set wordDoc = Nothing
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
    On Error GoTo 0
    Err.Raise Err.Number, "CrearDocumentoPrueba", "Error creando documento: " & Err.Description
End Sub
