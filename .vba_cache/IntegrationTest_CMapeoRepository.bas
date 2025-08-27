Attribute VB_Name = "IntegrationTest_CMapeoRepository"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CMapeoRepository
' Arquitectura: Pruebas Aisladas con Mocks
' ============================================================================

Private m_repository As IMapeoRepository
Private m_mockConfig As CMockConfig

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function IntegrationTest_CMapeoRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CMapeoRepository - Pruebas Unitarias CMapeoRepository"
    
    suiteResult.AddTestResult Test_GetMapeoPorTipo_Success()
    suiteResult.AddTestResult Test_GetMapeoPorTipo_NotFound()
    
    Set IntegrationTest_CMapeoRepository_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_mockConfig = New CMockConfig
    
    Dim repoImpl As New CMapeoRepository
    repoImpl.Initialize m_mockConfig
    Set m_repository = repoImpl
End Sub

Private Sub Teardown()
    Set m_repository = Nothing
    Set m_mockConfig = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function Test_GetMapeoPorTipo_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetMapeoPorTipo debe devolver un recordset con datos"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim tipoSolicitud As String
    tipoSolicitud = "PC"
    
    ' Configurar mockConfig para devolver una ruta de BD válida
    m_mockConfig.SetDataPath "C:\Test\Backend.accdb"
    m_mockConfig.SetDatabasePassword "testpass"
    
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
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_GetMapeoPorTipo_Success = testResult
End Function

Private Function Test_GetMapeoPorTipo_NotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetMapeoPorTipo debe devolver un recordset vacío si no hay mapeo"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Arrange
    Dim tipoSolicitud As String
    tipoSolicitud = "TIPO_INEXISTENTE"
    
    ' Configurar mockConfig para devolver una ruta de BD válida
    m_mockConfig.SetDataPath "C:\Test\Backend.accdb"
    m_mockConfig.SetDatabasePassword "testpass"
    
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
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_GetMapeoPorTipo_NotFound = testResult
End Function
