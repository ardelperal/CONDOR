Attribute VB_Name = "Test_CMapeoRepository"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CMapeoRepository
' Arquitectura: Pruebas Aisladas con Mocks
' ============================================================================

Private m_repository As IMapeoRepository
Private m_mockConfig As CMockConfig
Private m_mockErrorHandler As CMockErrorHandlerService

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function Test_CMapeoRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("Test_CMapeoRepository - Pruebas Unitarias CMapeoRepository")
    
    Call suiteResult.AddTestResult(Test_GetMapeoPorTipo_Success())
    Call suiteResult.AddTestResult(Test_GetMapeoPorTipo_NotFound())
    
    Set Test_CMapeoRepository_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_mockConfig = New CMockConfig
    Set m_mockErrorHandler = New CMockErrorHandlerService
    
    Dim repoImpl As New CMapeoRepository
    Call repoImpl.Initialize(m_mockConfig, m_mockErrorHandler)
    Set m_repository = repoImpl
End Sub

Private Sub Teardown()
    Set m_repository = Nothing
    Set m_mockConfig = Nothing
    Set m_mockErrorHandler = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function Test_GetMapeoPorTipo_Success() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetMapeoPorTipo debe devolver un recordset con datos")
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim tipoSolicitud As String
    tipoSolicitud = "PC"
    
    ' Configurar mockConfig para devolver una ruta de BD válida
    m_mockConfig.SetSetting "DATABASE_PATH", "C:\Test\Backend.accdb"
    m_mockConfig.SetSetting "DB_PASSWORD", "testpass"
    
    ' Act
    Dim rs As DAO.recordset
    Set rs = m_repository.GetMapeoPorTipo(tipoSolicitud)
    
    ' Assert
    AssertNotNull rs, "El recordset no debe ser nulo"
    AssertTrue Not rs.EOF, "El recordset no debe estar vacío"
    
    rs.Close
    Set rs = Nothing
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    m_mockConfig.Reset
    m_mockErrorHandler.Reset
    Call Teardown
    Set Test_GetMapeoPorTipo_Success = testResult
End Function

Private Function Test_GetMapeoPorTipo_NotFound() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetMapeoPorTipo debe devolver un recordset vacío si no hay mapeo")
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim tipoSolicitud As String
    tipoSolicitud = "TIPO_INEXISTENTE"
    
    ' Configurar mockConfig para devolver una ruta de BD válida
    m_mockConfig.SetSetting "DATABASE_PATH", "C:\Test\Backend.accdb"
    m_mockConfig.SetSetting "DB_PASSWORD", "testpass"
    
    ' Act
    Dim rs As DAO.recordset
    Set rs = m_repository.GetMapeoPorTipo(tipoSolicitud)
    
    ' Assert
    AssertNotNull rs, "El recordset no debe ser nulo"
    AssertTrue rs.EOF, "El recordset debe estar vacío"
    
    rs.Close
    Set rs = Nothing
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("Error inesperado: " & Err.Description)
Cleanup:
    m_mockConfig.Reset
    m_mockErrorHandler.Reset
    Call Teardown
    Set Test_GetMapeoPorTipo_NotFound = testResult
End Function
