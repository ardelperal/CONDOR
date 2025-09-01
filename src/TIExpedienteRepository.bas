Attribute VB_Name = "TIExpedienteRepository"
Option Compare Database
Option Explicit


Private Const TEMPLATE_PATH As String = "back\test_db\templates\Expedientes_test_template.accdb"
Private Const ACTIVE_PATH As String = "back\test_db\active\Expedientes_itest.accdb"

Public Function TIExpedienteRepositoryRunAll() As CTestSuiteResult
    Set TIExpedienteRepositoryRunAll = New CTestSuiteResult
    TIExpedienteRepositoryRunAll.Initialize "TIExpedienteRepository"
    TIExpedienteRepositoryRunAll.AddResult TestObtenerExpedientePorId_IntegrationSuccess()
End Function

Private Sub Setup()
    modTestUtils.PrepareTestDatabase modTestUtils.GetProjectPath & TEMPLATE_PATH, modTestUtils.GetProjectPath & ACTIVE_PATH
    
    ' Insertar expediente de prueba
    Dim db As Object
    Set db = CreateObject("DAO.DBEngine.0").OpenDatabase(modTestUtils.GetProjectPath & ACTIVE_PATH)
    
    Dim sql As String
    sql = "INSERT INTO Expedientes (idExpediente, Nemotecnico, Titulo, ContratistaPrincipal) " & _
          "VALUES (1, 'EXP-2024-001', 'Proyecto de Prueba Alfa', 'Contratista Principal S.A.')"
    
    db.Execute sql
    db.Close
    Set db = Nothing
End Sub

Private Sub Teardown()
    On Error Resume Next
    Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFile modTestUtils.GetProjectPath & ACTIVE_PATH, True
End Sub

Private Function TestObtenerExpedientePorId_IntegrationSuccess() As CTestResult
    Set TestObtenerExpedientePorId_IntegrationSuccess = New CTestResult
    TestObtenerExpedientePorId_IntegrationSuccess.Initialize "Debe recuperar un expediente completo de la BD de prueba"
    
    Dim repo As IExpedienteRepository
    Dim config As IConfig
    Dim result As EExpediente
    
    On Error GoTo TestFail
    
    Call Setup
    
    ' Arrange
    Set config = modConfigFactory.CreateConfigService()
    config.SetSetting "ExpedientesDBPath", modTestUtils.GetProjectPath & ACTIVE_PATH
    
    Set repo = modRepositoryFactory.CreateExpedienteRepository(config)
    
    ' Act
    Set result = repo.ObtenerExpedientePorId(1)
    
    ' Assert - PRUEBA FORTALECIDA
    modAssert.AssertNotNull result, "El expediente recuperado no debe ser nulo."
    modAssert.AssertEquals 1, result.idExpediente, "El ID no coincide."
    modAssert.AssertEquals "EXP-2024-001", result.Nemotecnico, "El nemotécnico no coincide."
    modAssert.AssertEquals "Proyecto de Prueba Alfa", result.Titulo, "El título no coincide."
    modAssert.AssertEquals "Contratista Principal S.A.", result.ContratistaPrincipal, "El contratista no coincide."
    
    TestObtenerExpedientePorId_IntegrationSuccess.Pass
    GoTo Cleanup

TestFail:
    TestObtenerExpedientePorId_IntegrationSuccess.Fail "Error: " & Err.Description
    
Cleanup:
    Call Teardown
    Set result = Nothing
    Set repo = Nothing
    Set config = Nothing
End Function

