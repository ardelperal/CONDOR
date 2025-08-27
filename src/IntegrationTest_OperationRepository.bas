Attribute VB_Name = "IntegrationTest_OperationRepository"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO: IntegrationTest_OperationRepository
' DESCRIPCIÓN: Suite de pruebas de integración para COperationRepository
' ARQUITECTURA: Testing - Validación de persistencia real en base de datos
' ============================================================================

' Constantes para las rutas de base de datos
Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\CONDOR_integration_test.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function IntegrationTest_OperationRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("IntegrationTest_OperationRepository - Pruebas de Integración COperationRepository")
    
    Call suiteResult.AddTestResult(Test_SaveLog_Integration_Success())
    
    Set IntegrationTest_OperationRepository_RunAll = suiteResult
End Function

' ============================================================================
' CONFIGURACIÓN DEL ENTORNO DE PRUEBA
' ============================================================================

Private Sub Setup()
    On Error GoTo ErrorHandler
    
    ' Aprovisionar la base de datos de prueba antes de cada ejecución
    Dim fullTemplatePath As String
    Dim fullTestPath As String
    
    fullTemplatePath = modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH
    fullTestPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    
    modTestUtils.PrepareTestDatabase fullTemplatePath, fullTestPath
    
    Debug.Print "Setup completado: BD de prueba preparada"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Setup (" & Err.Number & "): " & Err.Description
    Err.Raise Err.Number, "IntegrationTest_OperationRepository.Setup", Err.Description
End Sub



' ============================================================================
' LIMPIEZA DEL ENTORNO DE PRUEBA
' ============================================================================

Private Sub Teardown()
    On Error GoTo ErrorHandler
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Eliminar la BD de prueba
    Dim testDbPath As String
    testDbPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    
    If fs.FileExists(testDbPath) Then
        fs.DeleteFile testDbPath
        Debug.Print "Teardown completado: BD de prueba eliminada"
    End If
    
    Set fs = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en Teardown: " & Err.Description
    ' No relanzar error en Teardown para evitar enmascarar errores de prueba
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function Test_SaveLog_Integration_Success() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("Test_SaveLog_Integration_Success - Debe guardar correctamente un log de operación")
    
    On Error GoTo ErrorHandler
    
    ' Preparar entorno
    Call Setup
    
    ' ARRANGE: Crear instancias reales
    Dim config As CConfig
    Dim errorHandler As CErrorHandlerService
    Dim repository As COperationRepository
    
    Set config = New CConfig
    Set errorHandler = New CErrorHandlerService
    Set repository = New COperationRepository
    
    ' Configurar CConfig para apuntar a la BD de prueba
    config.SetSetting "DATABASE_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    config.SetSetting "DB_PASSWORD", "" ' Asumiendo que la BD de prueba no tiene contraseña
    
    ' Inyectar dependencias
    Call repository.Initialize(config, errorHandler)
    
    ' ACT: Ejecutar la operación a probar
    Dim testOperationType As String
    Dim testEntityId As String
    Dim testDetails As String
    
    testOperationType = "TEST_OP"
    testEntityId = "ENTITY_123"
    testDetails = "Detalles de la prueba de integración."
    
    repository.SaveLog testOperationType, testEntityId, testDetails
    
    ' ASSERT: Verificar que el registro se guardó correctamente
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recordCount As Long
    
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, dbFailOnError, False)
    
    ' Verificar que hay exactamente 1 registro
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS RecordCount FROM Tb_Operaciones_Log")
    recordCount = rs!RecordCount
    rs.Close
    
    modAssert.AssertEquals 1, recordCount, "Debe haber exactamente 1 registro en Tb_Operaciones_Log"
    
    ' Verificar que los datos son correctos
    Set rs = db.OpenRecordset("SELECT * FROM Tb_Operaciones_Log")
    
    modAssert.AssertEquals testOperationType, rs!TipoOperacion, "TipoOperacion debe coincidir"
    modAssert.AssertEquals testEntityId, rs!IDEntidadAfectada, "IDEntidadAfectada debe coincidir"
    modAssert.AssertEquals testDetails, rs!Detalles, "Detalles debe coincidir"
    
    ' Verificar que se guardaron campos automáticos
    modAssert.AssertTrue Not IsNull(rs!FechaHora), "FechaHora no debe ser nulo"
    modAssert.AssertTrue Not IsNull(rs!Usuario), "Usuario no debe ser nulo"
    
    rs.Close
    db.Close
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Call Teardown
    
    Set Test_SaveLog_Integration_Success = testResult
End Function