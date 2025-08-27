Attribute VB_Name = "Test_CWordManager"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CWordManager
' Arquitectura: Pruebas de Integración Controlada con Word
' ============================================================================

Private m_wordManager As IWordManager
Private m_mockConfig As CMockConfig
Private m_mockLogger As CMockOperationLogger
Private m_mockErrorHandler As CMockErrorHandlerService ' Added
Private m_tempDocPath As String
Private m_tempTemplatePath As String

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function Test_CWordManager_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CWordManager - Pruebas Unitarias CWordManager"
    
    suiteResult.AddTestResult Test_AbrirCerrarDocumento_Success()
    suiteResult.AddTestResult Test_ReemplazarTexto_Success()
    suiteResult.AddTestResult Test_GuardarDocumento_Success()
    
    Set Test_CWordManager_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_mockConfig = New CMockConfig
    Set m_mockLogger = New CMockOperationLogger
    Set m_mockErrorHandler = New CMockErrorHandlerService ' Instantiated
    
    ' Configurar mocks para rutas de ficheros temporales
    m_mockConfig.AddSetting "GENERATED_DOCS_PATH", Environ("TEMP") & "\CondorTests"
    m_mockConfig.AddSetting "PLANTILLA_PATH", Environ("TEMP") & "\CondorTests"
    
    ' Asegurarse de que la carpeta de tests existe
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        fso.CreateFolder Environ("TEMP") & "\CondorTests"
    End If
    
    ' Crear un documento de Word temporal para las pruebas
    m_tempTemplatePath = Environ("TEMP") & "\CondorTests\TestTemplate.docx"
    CreateDummyWordDocument m_tempTemplatePath, "[MARCADOR_TEST]"
    
    Dim managerImpl As New CWordManager
    managerImpl.Initialize m_mockErrorHandler ' Injected
    Set m_wordManager = managerImpl
End Sub

Private Sub Teardown()
    On Error Resume Next
    m_wordManager.CerrarDocumento ' Asegurarse de cerrar Word
    Set m_wordManager = Nothing
    Set m_mockConfig = Nothing
    Set m_mockLogger = Nothing
    Set m_mockErrorHandler = Nothing ' Cleaned up
    
    ' Limpiar ficheros temporales
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(m_tempTemplatePath) Then fso.DeleteFile m_tempTemplatePath
    If m_tempDocPath <> "" And fso.FileExists(m_tempDocPath) Then fso.DeleteFile m_tempDocPath
    
    ' Limpiar carpeta de tests si está vacía
    If fso.FolderExists(Environ("TEMP") & "\CondorTests") Then
        If fso.GetFolder(Environ("TEMP") & "\CondorTests").Files.Count = 0 And _
           fso.GetFolder(Environ("TEMP") & "\CondorTests").SubFolders.Count = 0 Then
            fso.DeleteFolder Environ("TEMP") & "\CondorTests"
        End If
    End If
    On Error GoTo 0
End Sub

' Helper para crear un documento Word dummy
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
    wordDoc.SaveAs2 filePath
    wordDoc.Close
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function Test_AbrirCerrarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "AbrirDocumento y CerrarDocumento deben funcionar correctamente"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Act
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    
    ' Assert
    AssertTrue opened, "El documento debe abrirse correctamente"
    
    m_wordManager.CerrarDocumento
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_AbrirCerrarDocumento_Success = testResult
End Function

Private Function Test_ReemplazarTexto_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ReemplazarTexto debe sustituir el marcador correctamente"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    Dim oldText As String
    oldText = "[MARCADOR_TEST]"
    Dim newText As String
    newText = "Texto Reemplazado"
    
    ' Act
    Dim replaced As Boolean
    replaced = m_wordManager.ReemplazarTexto(oldText, newText)
    
    ' Assert
    AssertTrue replaced, "El texto debe ser reemplazado"
    
    ' Verificar el contenido guardando y releyendo el documento
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
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_ReemplazarTexto_Success = testResult
End Function

Private Function Test_GuardarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GuardarDocumento debe guardar el documento en la ruta especificada"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim opened As Boolean
    opened = m_wordManager.AbrirDocumento(m_tempTemplatePath)
    AssertTrue opened, "Precondición: El documento debe abrirse correctamente"
    
    m_tempDocPath = Environ("TEMP") & "\CondorTests\SavedDoc.docx"
    
    ' Act
    Dim saved As Boolean
    saved = m_wordManager.GuardarDocumento(m_tempDocPath)
    
    ' Assert
    AssertTrue saved, "El documento debe guardarse correctamente"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    AssertTrue fso.FileExists(m_tempDocPath), "El archivo guardado debe existir"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_GuardarDocumento_Success = testResult
End Function
