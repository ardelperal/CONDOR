Attribute VB_Name = "TIOperationRepository"
Option Compare Database
Option Explicit


' ============================================================================
' MÓDULO: IntegrationTestOperationRepository
' DESCRIPCIÓN: Suite de pruebas de integración para COperationRepository
' ARQUITECTURA: Testing - Validación de persistencia real en base de datos
' ============================================================================

' Constantes para las rutas de base de datos
Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_integration_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function TIOperationRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TIOperationRepository")
    
    Call suiteResult.AddResult(TestSaveLogIntegrationSuccess())
    
    Set TIOperationRepositoryRunAll = suiteResult
End Function

' ============================================================================
' CONFIGURACIÓN DEL ENTORNO DE PRUEBA
' ============================================================================

Private Sub Setup()
    On Error GoTo TestError
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Exit Sub
TestError:
    Err.Raise Err.Number, "IntegrationTestOperationRepository.Setup", Err.Description
End Sub

' ============================================================================
' LIMPIEZA DEL ENTORNO DE PRUEBA
' ============================================================================

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testDbPath As String
    testDbPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    If fs.FileExists(testDbPath) Then
        fs.DeleteFile testDbPath
    End If
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestSaveLogIntegrationSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("SaveLog debe guardar correctamente un log de operación en la base de datos")
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo TestError
    
    Call Setup
    
    ' ARRANGE: Crear dependencias con factorías
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    config.SetSetting "DB_PASSWORD", ""
    
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    Dim repository As IOperationRepository
    Set repository = modRepositoryFactory.CreateOperationRepository()
    
    ' ACT: Ejecutar la operación a probar
    Dim testOperationType As String
    Dim testEntityId As String
    Dim testDetails As String
    testOperationType = "TEST_OP"
    testEntityId = "ENTITY_123"
    testDetails = "Detalles de la prueba de integración."
    
    repository.SaveLog testOperationType, testEntityId, testDetails
    
    ' ASSERT: Verificar que el registro se guardó correctamente
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, False, False)
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS RecordCount FROM Tb_Operaciones_Log")
    modAssert.AssertEquals 1, rs!recordCount, "Debe haber exactamente 1 registro en Tb_Operaciones_Log"
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT * FROM Tb_Operaciones_Log")
    modAssert.AssertEquals testOperationType, rs!tipoOperacion, "TipoOperacion debe coincidir"
    modAssert.AssertEquals testEntityId, rs!idEntidadAfectada, "IDEntidadAfectada debe coincidir"
    modAssert.AssertEquals testDetails, rs!detalles, "Detalles debe coincidir"
    modAssert.AssertTrue Not IsNull(rs!fechaHora), "FechaHora no debe ser nulo"
    modAssert.AssertTrue Not IsNull(rs!usuario), "Usuario no debe ser nulo"
    
    testResult.Pass
    GoTo Cleanup
    
TestError:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Call Teardown
    Set TestSaveLogIntegrationSuccess = testResult
End Function


