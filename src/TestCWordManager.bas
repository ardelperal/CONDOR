Attribute VB_Name = "TestCWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CWordManager
' Arquitectura: Pruebas de Integración Controlada con Word
' ============================================================================



' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function TestCWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestCWordManager")
    
    Call suiteResult.AddTestResult(TestAbrirCerrarDocumentoSuccess())
    Call suiteResult.AddTestResult(TestReemplazarTextoSuccess())
    Call suiteResult.AddTestResult(TestGuardarDocumentoSuccess())
    
    Set TestCWordManagerRunAll = suiteResult
End Function



Private Sub CreateDummyWordDocument(ByVal filePath As String, ByVal content As String)
    Dim wordApp As Object
    Dim wordDoc As Object
    On Error Resume Next
    Set wordApp = GetObject("Word.Application")
    If Err.Number <> 0 Then
        Set wordApp = CreateObject("Word.Application")
    End If
    Err.Clear
    
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    
    Set wordDoc = wordApp.Documents.Add
    wordDoc.Content.Text = content
    Call wordDoc.SaveAs2(filePath)
    wordDoc.Close
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function TestAbrirCerrarDocumentoSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("AbrirDocumento y CerrarDocumento deben funcionar correctamente")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks y variables locales
    Dim wordManager As IWordManager
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim mockLogger As New CMockOperationLogger
    mockLogger.Reset
    Dim mockErrorHandler As New CMockErrorHandlerService
    mockErrorHandler.Reset
    Dim tempDocPath As String
    Dim tempTemplatePath As String
    
    Call mockConfig.AddSetting("GENERATED_DOCS_PATH", Environ("TEMP") & "\CondorTests")
    Call mockConfig.AddSetting("PLANTILLA_PATH", Environ("TEMP") & "\CondorTests")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        Call fso.CreateFolder(Environ("TEMP") & "\CondorTests")
    End If
    
    tempTemplatePath = Environ("TEMP") & "\CondorTests\TestTemplate.docx"
    CreateDummyWordDocument tempTemplatePath, "[MARCADOR_TEST]"
    
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Dim managerImpl As IWordManager
    Set managerImpl = New CMockWordManager
    Call managerImpl.Initialize(wordApp, mockErrorHandler)
    Set wordManager = managerImpl
    
    ' Act - Ejecutar el método bajo prueba
    Dim opened As Boolean
    opened = wordManager.AbrirDocumento(tempTemplatePath)
    
    ' Assert - Verificar resultados
    AssertTrue opened, "El documento debe abrirse correctamente"
    
    wordManager.CerrarDocumento
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    ' Limpiar recursos
    On Error Resume Next
    If Not wordManager Is Nothing Then Call wordManager.CerrarDocumento()
    Set wordManager = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set fso = Nothing
    Set wordApp = Nothing
    Set managerImpl = Nothing
    
    Dim fs As IFileSystem
    Set fs = ModFileSystemFactory.CreateFileSystem()
    
    If tempTemplatePath <> "" And fs.FileExists(tempTemplatePath) Then fs.DeleteFile tempTemplatePath
    If tempDocPath <> "" And fs.FileExists(tempDocPath) Then fs.DeleteFile tempDocPath
    
    Dim testFolderPath As String
    testFolderPath = Environ("TEMP") & "\CondorTests"
    
    If fs.FolderExists(testFolderPath) Then
        Dim fsoCleanup As Object
        Set fsoCleanup = CreateObject("Scripting.FileSystemObject")
        If fsoCleanup.GetFolder(testFolderPath).Files.Count = 0 And _
           fsoCleanup.GetFolder(testFolderPath).SubFolders.Count = 0 Then
            fs.DeleteFolder testFolderPath
        End If
        Set fsoCleanup = Nothing
    End If
    
    Set fs = Nothing
    On Error GoTo 0
    Set TestAbrirCerrarDocumentoSuccess = testResult
End Function

Private Function TestReemplazarTextoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ReemplazarTexto debe sustituir el marcador correctamente"
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks y variables locales
    Dim wordManager As IWordManager
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockErrorHandler As New CMockErrorHandlerService
    Dim tempDocPath As String
    Dim tempTemplatePath As String
    
    Call mockConfig.AddSetting("GENERATED_DOCS_PATH", Environ("TEMP") & "\CondorTests")
    Call mockConfig.AddSetting("PLANTILLA_PATH", Environ("TEMP") & "\CondorTests")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        Call fso.CreateFolder(Environ("TEMP") & "\CondorTests")
    End If
    
    tempTemplatePath = Environ("TEMP") & "\CondorTests\TestTemplate.docx"
    CreateDummyWordDocument tempTemplatePath, "[MARCADOR_TEST]"
    
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Dim managerImpl As IWordManager
    Set managerImpl = New CMockWordManager
    Call managerImpl.Initialize(wordApp, mockErrorHandler)
    Set wordManager = managerImpl
    
    ' Act - Ejecutar el método bajo prueba
    Dim opened As Boolean
    opened = wordManager.AbrirDocumento(tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    Dim oldText As String
    oldText = "[MARCADOR_TEST]"
    Dim newText As String
    newText = "Texto Reemplazado"
    
    Dim replaced As Boolean
    replaced = wordManager.ReemplazarTexto(oldText, newText)
    
    ' Assert - Verificar resultados
    AssertTrue replaced, "El texto debe ser reemplazado"
    
    tempDocPath = Environ("TEMP") & "\CondorTests\ReplacedDoc.docx"
    Dim saved As Boolean
    saved = wordManager.GuardarDocumento(tempDocPath)
    AssertTrue saved, "El documento debe guardarse correctamente"
    
    wordManager.CerrarDocumento
    
    Dim fsoVerify As Object
    Set fsoVerify = CreateObject("Scripting.FileSystemObject")
    If fsoVerify.FileExists(tempDocPath) Then
        Dim wordAppVerify As Object
        Dim wordDoc As Object
        On Error Resume Next
        Set wordAppVerify = GetObject("Word.Application")
        If Err.Number <> 0 Then
            Set wordAppVerify = CreateObject("Word.Application")
        End If
        Err.Clear
        
        wordAppVerify.Visible = False
        wordAppVerify.DisplayAlerts = False
        
        Set wordDoc = wordAppVerify.Documents.Open(tempDocPath)
        Dim docContent As String
        docContent = wordDoc.Content.Text
        wordDoc.Close
        wordAppVerify.Quit
        Set wordDoc = Nothing
        Set wordAppVerify = Nothing
        
        AssertTrue InStr(docContent, newText) > 0, "El nuevo texto debe estar en el documento"
        AssertTrue InStr(docContent, oldText) = 0, "El marcador original no debe estar en el documento"
    Else
        testResult.Fail "El documento guardado no se encontró para verificación."
    End If
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    ' Limpiar recursos
    On Error Resume Next
    If Not wordManager Is Nothing Then Call wordManager.CerrarDocumento()
    Set wordManager = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set fso = Nothing
    Set fsoVerify = Nothing
    Set wordApp = Nothing
    Set managerImpl = Nothing
    
    Dim fs As IFileSystem
    Set fs = ModFileSystemFactory.CreateFileSystem()
    
    If tempTemplatePath <> "" And fs.FileExists(tempTemplatePath) Then fs.DeleteFile tempTemplatePath
    If tempDocPath <> "" And fs.FileExists(tempDocPath) Then fs.DeleteFile tempDocPath
    
    Dim testFolderPath As String
    testFolderPath = Environ("TEMP") & "\CondorTests"
    
    If fs.FolderExists(testFolderPath) Then
        Dim fsoCleanup As Object
        Set fsoCleanup = CreateObject("Scripting.FileSystemObject")
        If fsoCleanup.GetFolder(testFolderPath).Files.Count = 0 And _
           fsoCleanup.GetFolder(testFolderPath).SubFolders.Count = 0 Then
            fs.DeleteFolder testFolderPath
        End If
        Set fsoCleanup = Nothing
    End If
    
    Set fs = Nothing
    On Error GoTo 0
    Set TestReemplazarTextoSuccess = testResult
End Function

Private Function TestGuardarDocumentoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GuardarDocumento debe guardar el documento en la ruta especificada"
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks y variables locales
    Dim wordManager As IWordManager
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockErrorHandler As New CMockErrorHandlerService
    Dim tempDocPath As String
    Dim tempTemplatePath As String
    
    Call mockConfig.AddSetting("GENERATED_DOCS_PATH", Environ("TEMP") & "\CondorTests")
    Call mockConfig.AddSetting("PLANTILLA_PATH", Environ("TEMP") & "\CondorTests")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        Call fso.CreateFolder(Environ("TEMP") & "\CondorTests")
    End If
    
    tempTemplatePath = Environ("TEMP") & "\CondorTests\TestTemplate.docx"
    CreateDummyWordDocument tempTemplatePath, "[MARCADOR_TEST]"
    
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Dim managerImpl As IWordManager
    Set managerImpl = New CMockWordManager
    Call managerImpl.Initialize(wordApp, mockErrorHandler)
    Set wordManager = managerImpl
    
    ' Act - Ejecutar el método bajo prueba
    Dim opened As Boolean
    opened = wordManager.AbrirDocumento(tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    tempDocPath = Environ("TEMP") & "\CondorTests\SavedDoc.docx"
    
    Dim saved As Boolean
    saved = wordManager.GuardarDocumento(tempDocPath)
    
    ' Assert - Verificar resultados
    AssertTrue saved, "El documento debe guardarse correctamente"
    
    Dim fsoVerify As Object
    Set fsoVerify = CreateObject("Scripting.FileSystemObject")
    AssertTrue fsoVerify.FileExists(tempDocPath), "El archivo guardado debe existir"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    ' Limpiar recursos
    On Error Resume Next
    If Not wordManager Is Nothing Then Call wordManager.CerrarDocumento()
    Set wordManager = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set fso = Nothing
    Set fsoVerify = Nothing
    Set wordApp = Nothing
    Set managerImpl = Nothing
    
    Dim fs As IFileSystem
    Set fs = ModFileSystemFactory.CreateFileSystem()
    
    If tempTemplatePath <> "" And fs.FileExists(tempTemplatePath) Then fs.DeleteFile tempTemplatePath
    If tempDocPath <> "" And fs.FileExists(tempDocPath) Then fs.DeleteFile tempDocPath
    
    Dim testFolderPath As String
    testFolderPath = Environ("TEMP") & "\CondorTests"
    
    If fs.FolderExists(testFolderPath) Then
        Dim fsoCleanup As Object
        Set fsoCleanup = CreateObject("Scripting.FileSystemObject")
        If fsoCleanup.GetFolder(testFolderPath).Files.Count = 0 And _
           fsoCleanup.GetFolder(testFolderPath).SubFolders.Count = 0 Then
            fs.DeleteFolder testFolderPath
        End If
        Set fsoCleanup = Nothing
    End If
    
    Set fs = Nothing
    On Error GoTo 0
    Set TestGuardarDocumentoSuccess = testResult
End Function