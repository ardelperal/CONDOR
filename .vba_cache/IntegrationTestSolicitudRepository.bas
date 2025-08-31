Option Compare Database
Option Explicit

' ============================================================================
' PRUEBAS DE INTEGRACIÓN PARA CSolicitudRepository
' ============================================================================
' Este módulo contiene pruebas de integración que validan el comportamiento
' del repositorio CSolicitudRepository contra una base de datos real.
' ============================================================================

' Constantes para las rutas de base de datos de prueba
Private Const CONDOR_TEMPLATE_PATH As String = "back\test_db\templates\CONDOR_test_template.accdb"
Private Const CONDOR_ACTIVE_PATH As String = "back\test_db\active\CONDOR_integration_test.accdb"

' Función principal que ejecuta todas las pruebas del módulo
Public Function IntegrationTestSolicitudRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTestSolicitudRepository"
    
    ' Ejecutar todas las pruebas de integración
    suiteResult.AddTestResult TestGetSolicitudByIdSuccess()
    suiteResult.AddTestResult TestGetSolicitudByIdNotFound()
    suiteResult.AddTestResult TestSaveSolicitudNew()
    suiteResult.AddTestResult TestSaveSolicitudUpdate()
    suiteResult.AddTestResult TestExecuteQuery()
    suiteResult.AddTestResult TestCargarDatosEspecificosPC()
    suiteResult.AddTestResult TestCargarDatosEspecificosCDCA()
    suiteResult.AddTestResult TestCargarDatosEspecificosCDCASUB()
    
    Set IntegrationTestSolicitudRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS DE CONFIGURACIÓN Y LIMPIEZA
' ============================================================================

Private Sub Setup()
    On Error GoTo TestError

    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath() & CONDOR_TEMPLATE_PATH, modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado) VALUES (1, 'EXP-TEST-001', Now(), 'Pendiente')"
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado, tipoSolicitud) VALUES (2, 'EXP-PC-001', Now(), 'Pendiente', 'PC')"
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado, tipoSolicitud) VALUES (3, 'EXP-CDCA-001', Now(), 'Pendiente', 'CDCA')"
    db.Execute "INSERT INTO T_Solicitudes (idSolicitud, idExpediente, fechaCreacion, estado, tipoSolicitud) VALUES (4, 'EXP-CDCASUB-001', Now(), 'Pendiente', 'CDCASUB')"
    db.Execute "INSERT INTO TbDatos_PC (idSolicitud, refSuministrador, numPlanoEspecificacion) VALUES (2, 'REF-SUM-PC', 'PLANO-PC-001')"
    db.Execute "INSERT INTO TbDatos_CD_CA (idSolicitud, refSuministrador, numContrato) VALUES (3, 'REF-SUM-CDCA', 'CONTRATO-CDCA-002')"
    db.Execute "INSERT INTO TbDatos_CD_CA_SUB (idSolicitud, refSuministrador, refSubSuministrador) VALUES (4, 'REF-SUM-SUB', 'REF-SUB-SUM-003')"
    
    db.Close
    Set db = Nothing
    
    Exit Sub

TestError:
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    Err.Raise Err.Number, "IntegrationTestSolicitudRepository.Setup", Err.Description
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    Dim fullTestPath As String
    fullTestPath = modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH
    If fs.FileExists(fullTestPath) Then
        fs.DeleteFile fullTestPath
    End If
    Set fs = Nothing
End Sub

' ============================================================================
' PRUEBAS DE GetSolicitudById
' ============================================================================

Private Function TestGetSolicitudByIdSuccess() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetSolicitudById debe obtener una solicitud correctamente"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    
    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim result As Solicitud
    Set result = repository.GetSolicitudById(1)
    
    modAssert.AssertNotNull result, "El resultado no debería ser nulo."
    modAssert.AssertEquals 1, result.idSolicitud, "El ID de la solicitud no coincide."
    modAssert.AssertEquals "EXP-TEST-001", result.idExpediente, "El ID del expediente no coincide."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function TestGetSolicitudByIdNotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetSolicitudById debe devolver Nothing si la solicitud no existe"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim result As Solicitud
    Set result = repository.GetSolicitudById(999)
    
    modAssert.AssertIsNull result, "El resultado debería ser nulo para un ID inexistente."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE SaveSolicitud
' ============================================================================

Private Function TestSaveSolicitudNew() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud debe insertar una nueva solicitud correctamente"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim nuevaSolicitud As New Solicitud
    nuevaSolicitud.idSolicitud = 0
    nuevaSolicitud.idExpediente = "EXP-NEW-001"
    nuevaSolicitud.fechaCreacion = Date
    nuevaSolicitud.idEstadoInterno = 1
    
    Dim newId As Long
    newId = repository.SaveSolicitud(nuevaSolicitud)
    
    modAssert.AssertTrue newId > 1, "El nuevo ID debe ser mayor que 1."
    
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    Set rs = db.OpenRecordset("SELECT * FROM T_Solicitudes WHERE idSolicitud = " & newId)
    
    modAssert.AssertFalse rs.EOF, "El nuevo registro debe existir en la base de datos."
    modAssert.AssertEquals "EXP-NEW-001", rs("idExpediente").Value, "El ID del expediente debe coincidir."
    
    testResult.Pass
    
Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function TestSaveSolicitudUpdate() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "SaveSolicitud debe actualizar una solicitud existente correctamente"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim solicitud As Solicitud
    Set solicitud = repository.GetSolicitudById(1)
    
    solicitud.idExpediente = "EXP-TEST-UPDATED"
    
    Dim updatedId As Long
    updatedId = repository.SaveSolicitud(solicitud)
    
    modAssert.AssertEquals 1, updatedId, "El ID devuelto debe ser el mismo para actualización."
    
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH)
    Set rs = db.OpenRecordset("SELECT * FROM T_Solicitudes WHERE idSolicitud = 1")
    
    modAssert.AssertFalse rs.EOF, "El registro debe existir en la base de datos."
    modAssert.AssertEquals "EXP-TEST-UPDATED", rs("idExpediente").Value, "El ID del expediente debe estar actualizado."
    
    testResult.Pass
    
Cleanup:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE ExecuteQuery
' ============================================================================

Private Function TestExecuteQuery() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ExecuteQuery debe ejecutar una consulta parametrizada y devolver un recordset"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService
    Dim rs As Object

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim queryName As String
    queryName = "GET_SOLICITUD_BY_ID"
    
    Dim params As New Collection
    Dim param1 As New QueryParameter
    param1.ParameterName = "idSolicitud"
    param1.ParameterValue = 1
    param1.DataType = dbLong
    params.Add param1
    
    Set rs = repository.ExecuteQuery(queryName, params)
    
    modAssert.AssertNotNull rs, "El recordset no debe ser nulo."
    modAssert.AssertFalse rs.EOF, "El recordset no debe estar vacío."
    modAssert.AssertEquals 1, rs!idSolicitud.Value, "El valor del campo no es el esperado."
    
    testResult.Pass
    
Cleanup:
    If Not rs Is Nothing Then rs.Close
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' ============================================================================
' PRUEBAS DE CargarDatosEspecificos
' ============================================================================

Private Function TestCargarDatosEspecificosPC() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos PC"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim solicitud As Solicitud
    Set solicitud = repository.GetSolicitudById(2)
    
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosPC, "Los datos PC no deberían ser nulos."
    modAssert.AssertEquals "REF-SUM-PC", solicitud.datosPC.refSuministrador, "El refSuministrador de datos PC debe coincidir."
    modAssert.AssertEquals "PLANO-PC-001", solicitud.datosPC.numPlanoEspecificacion, "El numPlanoEspecificacion de datos PC debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function TestCargarDatosEspecificosCDCA() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos CD_CA"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim solicitud As Solicitud
    Set solicitud = repository.GetSolicitudById(3)
    
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosCDCA, "Los datos CD_CA no deberían ser nulos."
    modAssert.AssertEquals "REF-SUM-CDCA", solicitud.datosCDCA.refSuministrador, "El refSuministrador de datos CD_CA debe coincidir."
    modAssert.AssertEquals "CONTRATO-CDCA-002", solicitud.datosCDCA.numContrato, "El numContrato de datos CD_CA debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

Private Function TestCargarDatosEspecificosCDCASUB() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "CargarDatosEspecificos debe mapear correctamente datos CD_CA_SUB"
    
    Dim repository As ISolicitudRepository
    Dim testConfig As IConfig
    Dim errorHandler As IErrorHandlerService

    On Error GoTo TestError
    
    Setup
    
    Dim settings As New Collection
    settings.Add modTestUtils.GetProjectPath() & CONDOR_ACTIVE_PATH, "DATABASE_PATH"
    settings.Add "", "DB_PASSWORD"
    Set testConfig = modConfigFactory.CreateConfigServiceFromCollection(settings)
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(testConfig)
    Set repository = modRepositoryFactory.CreateSolicitudRepository(testConfig, errorHandler)
    
    Dim solicitud As Solicitud
    Set solicitud = repository.GetSolicitudById(4)
    
    modAssert.AssertNotNull solicitud, "La solicitud no debería ser nula."
    modAssert.AssertNotNull solicitud.datosCDCASUB, "Los datos CD_CA_SUB no deberían ser nulos."
    modAssert.AssertEquals "REF-SUM-SUB", solicitud.datosCDCASUB.refSuministrador, "El refSuministrador de datos CD_CA_SUB debe coincidir."
    modAssert.AssertEquals "REF-SUB-SUM-003", solicitud.datosCDCASUB.refSubSuministrador, "El refSubSuministrador de datos CD_CA_SUB debe coincidir."
    
    testResult.Pass
    
Cleanup:
    Teardown
    Exit Function
    
TestError:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function