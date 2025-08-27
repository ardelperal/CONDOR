Attribute VB_Name = "IntegrationTest_WorkflowRepository"
'''
' Módulo: IntegrationTest_WorkflowRepository
' Propósito: Pruebas de integración para CWorkflowRepository
' Autor: Sistema CONDOR
' Fecha: 2024
' 
' Descripción:
' Este módulo contiene pruebas de integración que validan el comportamiento
' de CWorkflowRepository contra una base de datos real.
'''

Option Compare Database
Option Explicit

' Constantes para las pruebas
Private Const TEST_DB_PATH As String = "back\test_db\CONDOR_integration_test.accdb"
Private Const TEMPLATE_DB_PATH As String = "back\test_db\CONDOR_template.accdb"
Private Const TEST_TIPO_SOLICITUD As String = "PC"
Private Const TEST_ESTADO_INICIAL As String = "BORRADOR"
Private Const TEST_ESTADO_REVISION As String = "EN_REVISION"
Private Const TEST_USUARIO_TEST As String = "TEST_USER"
Private Const TEST_ROL_ADMIN As String = "ADMINISTRADOR"

' Variables globales para las pruebas
Private m_Repository As CWorkflowRepository
Private testConfig As CConfig
Private m_TestSolicitudID As Long
Private m_Config As CConfig
Private activeTestPath As String

'''
' Ejecuta todas las pruebas de integración del repositorio
'''
Public Function IntegrationTest_WorkflowRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_WorkflowRepository - Pruebas de Integración"
    
    On Error GoTo ErrorHandler
    
    ' Ejecutar pruebas individuales
    suiteResult.AddTestResult Test_IsValidTransition_Integration()
    suiteResult.AddTestResult Test_GetAvailableStates_Integration()
    suiteResult.AddTestResult Test_GetNextStates_Integration()
    suiteResult.AddTestResult Test_GetInitialState_Integration()
    suiteResult.AddTestResult Test_IsStateFinal_Integration()
    suiteResult.AddTestResult Test_RecordStateChange_Integration()
    suiteResult.AddTestResult Test_GetStateHistory_Integration()
    suiteResult.AddTestResult Test_HasTransitionPermission_Integration()
    suiteResult.AddTestResult Test_RequiresApproval_Integration()
    suiteResult.AddTestResult Test_GetTransitionRequiredRole_Integration()
    
    Set IntegrationTest_WorkflowRepository_RunAll = suiteResult
    Exit Function
    
ErrorHandler:
    Dim errorResult As New CTestResult
    errorResult.Initialize "IntegrationTest_WorkflowRepository_RunAll_Error"
    errorResult.Fail "Error en suite: " & Err.Number & " - " & Err.Description
    suiteResult.AddTestResult errorResult
    
    Call TeardownTestEnvironment
    Set IntegrationTest_WorkflowRepository_RunAll = suiteResult
End Function

'''
' Configura el entorno de pruebas con base de datos separada
'''
Private Sub Setup()
    On Error GoTo ErrorHandler
    
    ' Aprovisionar la base de datos de prueba antes de cada ejecución
    Dim fullTemplatePath As String
    Dim fullTestPath As String
    
    fullTemplatePath = modTestUtils.GetProjectPath() & TEMPLATE_DB_PATH
    fullTestPath = modTestUtils.GetProjectPath() & TEST_DB_PATH
    activeTestPath = fullTestPath
    
    modTestUtils.PrepareTestDatabase fullTemplatePath, fullTestPath
    
    ' Configurar objetos de prueba
    Set testConfig = New CConfig
    testConfig.SetSetting "DATABASE_PATH", fullTestPath
    Set m_Config = testConfig
    
    Set m_Repository = New CWorkflowRepository
    m_Repository.Initialize testConfig
    
    ' Insertar datos de prueba en las tablas
    Call InsertTestData
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en Setup (" & Err.Number & "): " & Err.Description
    Err.Raise Err.Number, "IntegrationTest_WorkflowRepository.Setup", Err.Description
End Sub



'''
' Inserta datos de prueba en las tablas
'''
Private Sub InsertTestData()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(testConfig.GetSetting("DATABASE_PATH"), dbFailOnError, False)
    
    ' Insertar estados de prueba
    db.Execute "INSERT INTO TbEstados (Codigo, Descripcion, TipoSolicitud, EsInicial, EsFinal) VALUES ('BORRADOR', 'Borrador', 'PC', True, False)"
    db.Execute "INSERT INTO TbEstados (Codigo, Descripcion, TipoSolicitud, EsInicial, EsFinal) VALUES ('EN_REVISION', 'En Revisión', 'PC', False, False)"
    db.Execute "INSERT INTO TbEstados (Codigo, Descripcion, TipoSolicitud, EsInicial, EsFinal) VALUES ('APROBADO', 'Aprobado', 'PC', False, True)"
    
    ' Insertar transiciones de prueba
    db.Execute "INSERT INTO TbTransiciones (TipoSolicitud, EstadoOrigen, EstadoDestino, RolRequerido, RequiereAprobacion) VALUES ('PC', 'BORRADOR', 'EN_REVISION', 'USUARIO', False)"
    db.Execute "INSERT INTO TbTransiciones (TipoSolicitud, EstadoOrigen, EstadoDestino, RolRequerido, RequiereAprobacion) VALUES ('PC', 'EN_REVISION', 'APROBADO', 'APROBADOR', True)"
    
    db.Close
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en InsertTestData: " & Err.Number & " - " & Err.Description
End Sub

'''
' Limpia el entorno de pruebas
'''
Private Sub TeardownTestEnvironment()
    On Error GoTo ErrorHandler
    
    Debug.Print "Limpiando entorno de pruebas..."
    
    ' Limpiar solicitud de prueba
    If m_TestSolicitudID > 0 Then
        Call CleanupTestSolicitud(m_TestSolicitudID)
    End If
    
    ' Eliminar archivo de base de datos activa usando IFileSystem
    If Len(activeTestPath) > 0 Then
        Dim fs As IFileSystem
        Set fs = modFileSystemFactory.CreateFileSystem()
        
        If fs.FileExists(activeTestPath) Then
            fs.DeleteFile activeTestPath
            Debug.Print "Base de datos de prueba eliminada: " & activeTestPath
        End If
        
        Set fs = Nothing
    End If
    
    ' Limpiar objetos
    Set m_Repository = Nothing
    Set m_Config = Nothing
    Set testConfig = Nothing
    activeTestPath = ""
    m_TestSolicitudID = 0
    
    Debug.Print "Entorno limpiado."
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en TeardownTestEnvironment: " & Err.Number & " - " & Err.Description
End Sub

'''  
' Crea una solicitud de prueba en la base de datos
'''
Private Function CreateTestSolicitud() As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & TEST_DB_PATH, dbFailOnError, False)
    Set rs = db.OpenRecordset("T_Solicitudes", dbOpenDynaset)
    
    rs.AddNew
    rs("TipoSolicitud") = TEST_TIPO_SOLICITUD
    rs("EstadoActual") = TEST_ESTADO_INICIAL
    rs("FechaCreacion") = Now()
    rs("UsuarioCreacion") = TEST_USUARIO_TEST
    rs("Descripcion") = "Solicitud de prueba para integración"
    rs("idExpediente") = "EXP-TEST-WF-001"
    rs.Update
    
    CreateTestSolicitud = rs("idSolicitud")
    
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
    CreateTestSolicitud = 0
    Debug.Print "ERROR en CreateTestSolicitud: " & Err.Number & " - " & Err.Description
End Function

'''
' Limpia la solicitud de prueba de la base de datos
'''
Private Sub CleanupTestSolicitud(ByVal solicitudID As Long)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    
    Set db = DBEngine.OpenDatabase(modTestUtils.GetProjectPath() & TEST_DB_PATH, dbFailOnError, False)
    
    ' Limpiar historial de estados
    db.Execute "DELETE FROM TbHistorialEstados WHERE idSolicitud = " & solicitudID
    
    ' Limpiar solicitud
    db.Execute "DELETE FROM T_Solicitudes WHERE idSolicitud = " & solicitudID
    
    db.Close
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en CleanupTestSolicitud: " & Err.Number & " - " & Err.Description
End Sub

'''
' Prueba de integración: IsValidTransition
'''
Private Function Test_IsValidTransition_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_IsValidTransition_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Arrange: Configurar datos de prueba
    ' (Los datos ya están insertados en Setup)
    
    ' Act & Assert: Probar transición válida
    Dim resultado As Boolean
    resultado = m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION)
    modAssert.AssertTrue resultado, "Transición válida de BORRADOR a EN_REVISION debe ser permitida"
    
    ' Act & Assert: Probar transición inválida
    resultado = m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, "APROBADO")
    modAssert.AssertFalse resultado, "Transición inválida de BORRADOR a APROBADO debe ser rechazada"
    
    testResult.Pass
    
Cleanup:
    Set Test_IsValidTransition_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: GetAvailableStates
'''
Private Function Test_GetAvailableStates_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetAvailableStates_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act: Obtener estados disponibles
    Dim estados As Collection
    Set estados = m_Repository.GetAvailableStates(TEST_TIPO_SOLICITUD)
    
    ' Assert: Verificar resultados
    modAssert.AssertNotNothing estados, "GetAvailableStates no debe devolver Nothing"
    modAssert.AssertTrue estados.Count > 0, "Debe haber al menos un estado disponible"
    modAssert.AssertTrue estados.Count = 3, "Deben existir exactamente 3 estados para PC"
    
    testResult.Pass
    
Cleanup:
    Set estados = Nothing
    Set Test_GetAvailableStates_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: GetNextStates
'''
Private Function Test_GetNextStates_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetNextStates_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act: Obtener estados siguientes desde BORRADOR
    Dim nextStates As Collection
    Set nextStates = m_Repository.GetNextStates(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL)
    
    ' Assert: Verificar resultados
    modAssert.AssertNotNothing nextStates, "GetNextStates no debe devolver Nothing"
    modAssert.AssertTrue nextStates.Count > 0, "Debe haber al menos un estado siguiente desde BORRADOR"
    modAssert.AssertTrue nextStates.Count = 1, "Desde BORRADOR solo debe haber 1 transición (a EN_REVISION)"
    
    testResult.Pass
    
Cleanup:
    Set nextStates = Nothing
    Set Test_GetNextStates_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: GetInitialState
'''
Private Function Test_GetInitialState_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetInitialState_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act: Obtener estado inicial
    Dim estadoInicial As String
    estadoInicial = m_Repository.GetInitialState(TEST_TIPO_SOLICITUD)
    
    ' Assert: Verificar resultados
    modAssert.AssertNotEmpty estadoInicial, "GetInitialState no debe devolver cadena vacía"
    modAssert.AssertEqual estadoInicial, TEST_ESTADO_INICIAL, "Estado inicial debe ser BORRADOR"
    
    testResult.Pass
    
Cleanup:
    Set Test_GetInitialState_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: IsStateFinal
'''
Private Function Test_IsStateFinal_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_IsStateFinal_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act & Assert: Probar estado no final (BORRADOR)
    Dim esFinal As Boolean
    esFinal = m_Repository.IsStateFinal(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL)
    modAssert.AssertFalse esFinal, "Estado BORRADOR no debe ser final"
    
    ' Act & Assert: Probar estado final (APROBADO)
    esFinal = m_Repository.IsStateFinal(TEST_TIPO_SOLICITUD, "APROBADO")
    modAssert.AssertTrue esFinal, "Estado APROBADO debe ser final"
    
    testResult.Pass
    
Cleanup:
    Set Test_IsStateFinal_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: RecordStateChange
'''
Private Function Test_RecordStateChange_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_RecordStateChange_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Arrange: Crear una solicitud de prueba
    Dim testSolicitudID As Long
    testSolicitudID = CreateTestSolicitud()
    
    ' Act: Registrar cambio de estado
    Dim resultado As Boolean
    resultado = m_Repository.RecordStateChange(testSolicitudID, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_USUARIO_TEST, "Cambio de prueba")
    
    ' Assert: Verificar que el cambio se registró correctamente
    modAssert.AssertTrue resultado, "El cambio de estado debe registrarse correctamente"
    
    testResult.Pass
    
Cleanup:
    ' Limpiar solicitud de prueba
    If testSolicitudID > 0 Then
        Call CleanupTestSolicitud(testSolicitudID)
    End If
    Set Test_RecordStateChange_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: GetStateHistory
'''
Private Function Test_GetStateHistory_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetStateHistory_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Arrange: Crear una solicitud de prueba y registrar cambio de estado
    Dim testSolicitudID As Long
    testSolicitudID = CreateTestSolicitud()
    m_Repository.RecordStateChange testSolicitudID, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_USUARIO_TEST, "Cambio de prueba"
    
    ' Act: Obtener historial
    Dim historial As Collection
    Set historial = m_Repository.GetStateHistory(testSolicitudID)
    
    ' Assert: Verificar resultados
    modAssert.AssertNotNothing historial, "GetStateHistory no debe devolver Nothing"
    modAssert.AssertTrue historial.Count > 0, "Debe existir al menos un registro en el historial"
    
    testResult.Pass
    
Cleanup:
    ' Limpiar solicitud de prueba
    If testSolicitudID > 0 Then
        Call CleanupTestSolicitud(testSolicitudID)
    End If
    Set historial = Nothing
    Set Test_GetStateHistory_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: HasTransitionPermission
'''
Private Function Test_HasTransitionPermission_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_HasTransitionPermission_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act & Assert: Probar permiso válido (USUARIO puede hacer BORRADOR -> EN_REVISION)
    Dim tienePermiso As Boolean
    tienePermiso = m_Repository.HasTransitionPermission(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, "USUARIO")
    modAssert.AssertTrue tienePermiso, "Usuario debe tener permiso para transición BORRADOR -> EN_REVISION"
    
    ' Act & Assert: Probar permiso inválido (rol inexistente)
    tienePermiso = m_Repository.HasTransitionPermission(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, "ROL_INEXISTENTE")
    modAssert.AssertFalse tienePermiso, "Rol inexistente no debe tener permisos"
    
    ' Act & Assert: Probar transición que requiere APROBADOR
    tienePermiso = m_Repository.HasTransitionPermission(TEST_TIPO_SOLICITUD, TEST_ESTADO_REVISION, "APROBADO", "APROBADOR")
    modAssert.AssertTrue tienePermiso, "Aprobador debe tener permiso para transición EN_REVISION -> APROBADO"
    
    testResult.Pass
    
Cleanup:
    Set Test_HasTransitionPermission_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: RequiresApproval
'''
Private Function Test_RequiresApproval_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_RequiresApproval_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act & Assert: Probar transición que requiere aprobación (EN_REVISION -> APROBADO)
    Dim requiereAprobacion As Boolean
    requiereAprobacion = m_Repository.RequiresApproval(TEST_TIPO_SOLICITUD, TEST_ESTADO_REVISION, "APROBADO")
    modAssert.AssertTrue requiereAprobacion, "Transición EN_REVISION -> APROBADO debe requerir aprobación"
    
    ' Act & Assert: Probar transición que no requiere aprobación (BORRADOR -> EN_REVISION)
    requiereAprobacion = m_Repository.RequiresApproval(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION)
    modAssert.AssertFalse requiereAprobacion, "Transición BORRADOR -> EN_REVISION no debe requerir aprobación"
    
    testResult.Pass
    
Cleanup:
    Set Test_RequiresApproval_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'''
' Prueba de integración: GetTransitionRequiredRole
'''
Private Function Test_GetTransitionRequiredRole_Integration() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetTransitionRequiredRole_Integration"
    
    On Error GoTo ErrorHandler
    
    ' Setup: Preparar base de datos de prueba
    Call Setup
    
    ' Act: Obtener rol requerido para transición BORRADOR -> EN_REVISION
    Dim rolRequerido As String
    rolRequerido = m_Repository.GetTransitionRequiredRole(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION)
    
    ' Assert: Verificar que se obtuvo el rol correcto
    modAssert.AssertNotEmpty rolRequerido, "Debe obtenerse un rol requerido para la transición"
    modAssert.AssertEqual rolRequerido, "USUARIO", "El rol requerido para BORRADOR -> EN_REVISION debe ser USUARIO"
    
    ' Act: Obtener rol requerido para transición EN_REVISION -> APROBADO
    rolRequerido = m_Repository.GetTransitionRequiredRole(TEST_TIPO_SOLICITUD, TEST_ESTADO_REVISION, "APROBADO")
    
    ' Assert: Verificar que se obtuvo el rol correcto
    modAssert.AssertNotEmpty rolRequerido, "Debe obtenerse un rol requerido para la transición"
    modAssert.AssertEqual rolRequerido, "APROBADOR", "El rol requerido para EN_REVISION -> APROBADO debe ser APROBADOR"
    
    testResult.Pass
    
Cleanup:
    Set Test_GetTransitionRequiredRole_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function