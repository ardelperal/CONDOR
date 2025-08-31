Attribute VB_Name = "IntegrationTestWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: IntegrationTestWordManager
' Descripción: Pruebas de integración para CWordManager
' ============================================================================

Private m_TempFolder As String
Private m_TempFiles As Collection

' ============================================================================
' CONFIGURACIÓN DE PRUEBAS
' ============================================================================

Public Function IntegrationTestWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestWordManager"
    
    suiteResult.AddTestResult IntegrationTestWordManagerCicloCompletoSuccess()
    suiteResult.AddTestResult IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse()
    
    Set IntegrationTestWordManagerRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function IntegrationTestWordManagerCicloCompletoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Ciclo completo de WordManager (Abrir, Reemplazar, Guardar, Cerrar) debe tener éxito"
    
    On Error GoTo TestError
    
    Setup
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim archivoOriginal As String
    Dim archivoGuardado As String
    Dim contenidoFinal As String
    
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    Set config = modConfigFactory.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    archivoOriginal = m_TempFolder & "documento_original.docx"
    CrearDocumentoPrueba archivoOriginal, "Hola [NOMBRE], este es un documento de prueba."
    m_TempFiles.Add archivoOriginal
    
    archivoGuardado = m_TempFolder & "documento_modificado.docx"
    m_TempFiles.Add archivoGuardado
    
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
    Teardown
    Set IntegrationTestWordManagerCicloCompletoSuccess = testResult
End Function

Private Function IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "AbrirDocumento con un fichero inexistente debe devolver False"
    
    On Error GoTo TestError
    
    Setup
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim rutaInvalida As String
    Dim resultado As Boolean
    
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    Set config = modConfigFactory.CreateConfigService()
    Set fileSystem = modFileSystemFactory.CreateFileSystem()
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    Set wordManager = modWordManagerFactory.CreateWordManager()
    
    rutaInvalida = m_TempFolder & "archivo_que_no_existe.docx"
    
    ' Act
    resultado = wordManager.AbrirDocumento(rutaInvalida)
    
    ' Assert
    modAssert.AssertFalse resultado, "Debería devolver False al intentar abrir un archivo inexistente"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    
Cleanup:
    Teardown
    Set IntegrationTestWordManagerAbrirFicheroInexistenteDevuelveFalse = testResult
End Function

' ============================================================================
' MÉTODOS DE SETUP Y TEARDOWN CENTRALIZADOS
' ============================================================================

Private Sub Setup()
    On Error GoTo TestError
    Set m_TempFiles = New Collection
    m_TempFolder = modTestUtils.GetProjectPath() & "back\test_env\word_tests\"
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    If Not fs.FolderExists(m_TempFolder) Then
        fs.CreateFolder m_TempFolder
    End If
    Set fs = Nothing
    Exit Sub
TestError:
    Err.Raise Err.Number, "Setup", "Error en inicialización: " & Err.Description
End Sub

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

Private Sub Teardown()
    On Error Resume Next
    Dim i As Integer
    Dim archivo As String
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    If Not m_TempFiles Is Nothing Then
        For i = 1 To m_TempFiles.Count
            archivo = m_TempFiles(i)
            If fs.FileExists(archivo) Then
                fs.DeleteFile archivo
            End If
        Next i
    End If
    If fs.FolderExists(m_TempFolder) Then
        fs.DeleteFolder m_TempFolder
    End If
    Set fs = Nothing
    On Error GoTo 0
End Sub
