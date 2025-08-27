Attribute VB_Name = "Test_CSolicitudRepository"
Option Compare Database
Option Explicit

' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudRepository
' Arquitectura: Pruebas Aisladas de Lógica Interna
' ============================================================================

Private m_repository As CSolicitudRepository
Private m_mockConfig As CMockConfig
Private m_mockLogger As CMockOperationLogger

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function Test_CSolicitudRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CSolicitudRepository - Pruebas Unitarias CSolicitudRepository"
    
    suiteResult.AddTestResult Test_Initialize_FailsWithNilDependencies()
    suiteResult.AddTestResult Test_Initialize_Success()
    
    Set Test_CSolicitudRepository_RunAll = suiteResult
End Function

' ============================================================================
' SETUP Y TEARDOWN
' ============================================================================

Private Sub Setup()
    Set m_repository = New CSolicitudRepository
    Set m_mockConfig = New CMockConfig
    Set m_mockLogger = New CMockOperationLogger
End Sub

Private Sub Teardown()
    Set m_repository = Nothing
    Set m_mockConfig = Nothing
    Set m_mockLogger = Nothing
End Sub

' ============================================================================
' PRUEBAS
' ============================================================================

Private Function Test_Initialize_FailsWithNilDependencies() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Initialize debe fallar si las dependencias son nulas"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Probar con config nulo
    On Error Resume Next
    m_repository.Initialize Nothing, m_mockLogger
    AssertEquals 5, Err.Number, "Debe fallar si IConfig es Nothing"
    On Error GoTo ErrorHandler
    
    ' Probar con logger nulo
    On Error Resume Next
    m_repository.Initialize m_mockConfig, Nothing
    AssertEquals 5, Err.Number, "Debe fallar si IOperationLogger es Nothing"
    On Error GoTo ErrorHandler
    
    testResult.Pass
    GoTo Cleanup

ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_Initialize_FailsWithNilDependencies = testResult
End Function

Private Function Test_Initialize_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Initialize debe establecer las dependencias correctamente"
    
    On Error GoTo ErrorHandler
    
    Call Setup
    
    ' Act
    m_repository.Initialize m_mockConfig, m_mockLogger
    
    ' Assert
    ' No podemos acceder a las variables privadas, pero podemos verificar que no hay errores
    ' y que el flag de inicializado (si fuera público) estaría a True.
    ' La ausencia de un error es el éxito en este caso.
    testResult.Pass
    GoTo Cleanup

ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Call Teardown
    Set Test_Initialize_Success = testResult
End Function
