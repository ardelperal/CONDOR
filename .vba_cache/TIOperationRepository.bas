Attribute VB_Name = "TIOperationRepository"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO: TIOperationRepository
' DESCRIPCIÓN: Suite de pruebas de integración para COperationRepository
' ARQUITECTURA: Patrón de Oro (Setup a Nivel de Suite + Transacciones)
' ============================================================================

Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_operation_itest.accdb"

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function TIOperationRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIOperationRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSaveLog_Success()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIOperationRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    Dim projectPath As String: projectPath = modTestUtils.GetProjectPath()
    Dim templatePath As String: templatePath = projectPath & CONDOR_TEMPLATE_PATH
    Dim activePath As String: activePath = projectPath & CONDOR_ACTIVE_PATH
    Call modTestUtils.SuiteSetup(templatePath, activePath)
End Sub

Private Sub SuiteTeardown()
    Dim activePath As String: activePath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Call modTestUtils.SuiteTeardown(activePath)
End Sub

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestSaveLog_Success() As CTestResult
    Set TestSaveLog_Success = New CTestResult
    TestSaveLog_Success.Initialize "SaveLog debe guardar correctamente un log de operación"
    
    Dim localConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim repository As IOperationRepository
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fs As IFileSystem
    Dim dbPath As String
    
    On Error GoTo TestFail

    ' Arrange: 1. Crear configuración local apuntando a la BD de prueba activa
    Dim mockConfigImpl As New CMockConfig
    mockConfigImpl.SetSetting "DATA_PATH", modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    Set localConfig = mockConfigImpl
    
    ' Arrange: 2. Crear dependencias inyectando la configuración local
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(localConfig)
    Set repository = modRepositoryFactory.CreateOperationRepository(localConfig, errorHandler)
    
    ' Arrange: Conectar a la base de datos activa de forma segura
    Set fs = modFileSystemFactory.CreateFileSystem()
    dbPath = localConfig.GetDataPath()
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 101, "Test.Arrange", "La BD de prueba de Operaciones no existe en la ruta esperada: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)

    ' Act: Ejecutar la operación a probar dentro de una transacción
    DBEngine.BeginTrans
    repository.SaveLog "TEST_OP", "ENTITY_123", "Detalles de prueba."
    
    ' Assert: Verificar directamente en la BD que el registro se insertó
    Set rs = db.OpenRecordset("SELECT * FROM tbOperacionesLog WHERE tipoOperacion = 'TEST_OP'")
    modAssert.AssertFalse rs.EOF, "Se debería haber insertado un registro de log."
    modAssert.AssertEquals "ENTITY_123", rs!idEntidad, "El ID de entidad no coincide."
    modAssert.AssertEquals "Detalles de prueba.", rs!detalles, "Los detalles no coinciden."

    TestSaveLog_Success.Pass
    GoTo Cleanup

TestFail:
    TestSaveLog_Success.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    DBEngine.Rollback
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Set repository = Nothing
    Set errorHandler = Nothing
    Set localConfig = Nothing
    Set fs = Nothing
End Function


