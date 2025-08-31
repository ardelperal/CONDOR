Attribute VB_Name = "TestCWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CWordManager
' Arquitectura: Pruebas de Integración Controlada con Word
' ============================================================================

Private m_wordManager As IWordManager
Private m_mockConfig As CMockConfig
Private m_mockLogger As CMockOperationLogger
Private m_mockErrorHandler As CMockErrorHandlerService
Private m_tempDocPath As String
Private m_tempTemplatePath As String

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

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_mockConfig = New CMockConfig
    Set m_mockLogger = New CMockOperationLogger
    Set m_mockErrorHandler = New CMockErrorHandlerService
    
    Call m_mockConfig.AddSetting("GENERATED_DOCS_PATH", Environ("TEMP") & "\CondorTests")
    Call m_mockConfig.AddSetting("PLANTILLA_PATH", Environ("TEMP") & "\CondorTests")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        Call fso.CreateFolder(Environ("TEMP") & "\CondorTests")
    End If
    
    m_tempTemplatePath = Environ("TEMP") & "\CondorTests\TestTemplate.docx"
    CreateDummyWordDocument m_tempTemplatePath, "[MARCADOR_TEST]"
    
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Dim managerImpl As New CWordManager
    Call managerImpl.Initialize(wordApp, m_mockErrorHandler)
    Set m_wordManager = managerImpl
End Sub

Private Sub Teardown()
    On Error Resume Next
    Call m_wordManager.CerrarDocumento()
    Set m_wordManager = Nothing
    Set m_mockConfig = Nothing
    Set m_mockLogger = Nothing
    Set m_mockErrorHandler = Nothing
    
    Dim fs As IFileSystem
    Set fs = ModFileSystemFactory.CreateFileSystem()
    
    If fs.FileExists(m_tempTemplatePath) Then fs.DeleteFile m_tempTemplatePath
    If m_tempDocPath <> "" And fs.FileExists(m_tempDocPath) Then fs.DeleteFile m_tempDocPath
    
    Dim testFolderPath As String
    testFolderPath = Environ("TEMP") & "\CondorTests"
    
    If fs.FolderExists(testFolderPath) Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.GetFolder(testFolderPath).Files.Count = 0 And _
           fso.GetFolder(testFolderPath).SubFolders.Count = 0 Then
            fs.DeleteFolder testFolderPath
        End If
        Set fso = Nothing
    End If
    
    Set fs = Nothing
    On Error GoTo 0
End Sub

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
    
    Call Setup
    
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    
    AssertTrue opened, "El documento debe abrirse correctamente"
    
    m_wordManager.CerrarDocumento
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    Call Teardown
    Set TestAbrirCerrarDocumentoSuccess = testResult
End Function

Private Function TestReemplazarTextoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ReemplazarTexto debe sustituir el marcador correctamente"
    
    On Error GoTo TestFail
    
    Call Setup
    
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    Dim oldText As String
    oldText = "[MARCADOR_TEST]"
    Dim newText As String
    newText = "Texto Reemplazado"
    
    Dim replaced As Boolean
    replaced = m_wordManager.ReemplazarTexto(oldText, newText)
    
    AssertTrue replaced, "El texto debe ser reemplazado"
    
    m_tempDocPath = Environ("TEMP") & "\CondorTests\ReplacedDoc.docx"
    Dim saved As Boolean
    saved = m_wordManager.GuardarDocumento(m_tempDocPath)
    AssertTrue saved, "El documento debe guardarse correctamente"
    
    m_wordManager.CerrarDocumento
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(m_tempDocPath) Then
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
        
        Set wordDoc = wordApp.Documents.Open(m_tempDocPath)
        Dim docContent As String
        docContent = wordDoc.Content.Text
        wordDoc.Close
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
        
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
    Call Teardown
    Set TestReemplazarTextoSuccess = testResult
End Function

Private Function TestGuardarDocumentoSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GuardarDocumento debe guardar el documento en la ruta especificada"
    
    On Error GoTo TestFail
    
    Call Setup
    
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    m_tempDocPath = Environ("TEMP") & "\CondorTests\SavedDoc.docx"
    
    Dim saved As Boolean
    saved = m_wordManager.GuardarDocumento(m_tempDocPath)
    
    AssertTrue saved, "El documento debe guardarse correctamente"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    AssertTrue fso.FileExists(m_tempDocPath), "El archivo guardado debe existir"
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set TestGuardarDocumentoSuccess = testResult
End Function