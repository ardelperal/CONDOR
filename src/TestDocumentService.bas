Option Compare Database
Option Explicit

Public Function TestDocumentServiceRunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "DocumentService"
    suite.AddResult TestGenerarDocumentoSuccess()
    Set TestDocumentServiceRunAll = suite
End Function

Private Function TestGenerarDocumentoSuccess() As CTestResult
    Set TestGenerarDocumentoSuccess = New CTestResult
    TestGenerarDocumentoSuccess.Initialize "GenerarDocumento con ID válido debe retornar una ruta"
    
    On Error GoTo TestFail
    
    ' Arrange
    Dim mockDocService As New CMockDocumentService
    mockDocService.Reset
    
    Dim docService As IDocumentService
    Set docService = mockDocService
    
    Dim expectedPath As String
    expectedPath = "C:\ruta\ficticia.docx"
    mockDocService.ConfigureGenerarDocumento expectedPath
    
    ' Act
    Dim resultPath As String
    resultPath = docService.GenerarDocumento(123)
    
    ' Assert
    AssertEquals expectedPath, resultPath, "La ruta devuelta no es la esperada."
    AssertTrue mockDocService.GenerarDocumento_WasCalled, "El método GenerarDocumento no fue llamado."
    AssertEquals 123, mockDocService.GenerarDocumento_LastSolicitudId, "El ID de solicitud no es el esperado."
    
    TestGenerarDocumentoSuccess.Pass
    GoTo Cleanup

TestFail:
    TestGenerarDocumentoSuccess.Fail "Error: " & Err.Description

Cleanup:
    Set docService = Nothing
    Set mockDocService = Nothing
End Function


