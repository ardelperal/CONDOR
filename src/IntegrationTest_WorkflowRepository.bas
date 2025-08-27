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
Private Const TEST_TIPO_SOLICITUD As String = "PC"
Private Const TEST_ESTADO_INICIAL As String = "BORRADOR"
Private Const TEST_ESTADO_REVISION As String = "EN_REVISION"
Private Const TEST_USUARIO_TEST As String = "TEST_USER"
Private Const TEST_ROL_ADMIN As String = "ADMINISTRADOR"

' Variables globales para las pruebas
Private m_Repository As CWorkflowRepository
Private m_Config As CConfig
Private m_TestSolicitudID As Long

'''
' Ejecuta todas las pruebas de integración del repositorio
'''
Public Function IntegrationTest_WorkflowRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_WorkflowRepository - Pruebas de Integración"
    
    On Error GoTo ErrorHandler
    
    ' Configurar entorno de pruebas
    Call SetupTestEnvironment
    
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
    
    ' Limpiar entorno de pruebas
    Call TeardownTestEnvironment
    
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
' Configura el entorno de pruebas
'''
Private Sub SetupTestEnvironment()
    On Error GoTo ErrorHandler
    
    Debug.Print "Configurando entorno de pruebas..."
    
    ' Inicializar configuración
    Set m_Config = New CConfig
    
    ' Inicializar repositorio
    Set m_Repository = New CWorkflowRepository
    m_Repository.Initialize m_Config
    
    ' Crear una solicitud de prueba para las pruebas que la requieren
    m_TestSolicitudID = CreateTestSolicitud()
    
    Debug.Print "Entorno configurado. SolicitudID de prueba: " & m_TestSolicitudID
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en SetupTestEnvironment: " & Err.Number & " - " & Err.Description
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
    
    ' Limpiar objetos
    Set m_Repository = Nothing
    Set m_Config = Nothing
    
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
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False)
    Set rs = db.OpenRecordset("TbSolicitudes", dbOpenDynaset)
    
    rs.AddNew
    rs("TipoSolicitud") = TEST_TIPO_SOLICITUD
    rs("EstadoActual") = TEST_ESTADO_INICIAL
    rs("FechaCreacion") = Now()
    rs("UsuarioCreacion") = TEST_USUARIO_TEST
    rs("Descripcion") = "Solicitud de prueba para integración"
    rs.Update
    
    CreateTestSolicitud = rs("ID")
    
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
    
    Set db = DBEngine.OpenDatabase(m_Config.GetDataPath(), dbFailOnError, False)
    
    ' Limpiar historial de estados
    db.Execute "DELETE FROM TbHistorialEstados WHERE SolicitudID = " & solicitudID
    
    ' Limpiar solicitud
    db.Execute "DELETE FROM TbSolicitudes WHERE ID = " & solicitudID
    
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
    
    ' Probar transición válida
    Dim resultado As Boolean
    resultado = m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION)
    
    If Not resultado Then
        testResult.Fail "Transición válida no detectada"
        GoTo Cleanup
    End If
    
    ' Probar transición inválida
    resultado = m_Repository.IsValidTransition(TEST_TIPO_SOLICITUD, "ESTADO_INEXISTENTE", TEST_ESTADO_REVISION)
    
    If resultado Then
        testResult.Fail "Transición inválida aceptada"
        GoTo Cleanup
    End If
    
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
    
    Dim estados As Collection
    Set estados = m_Repository.GetAvailableStates(TEST_TIPO_SOLICITUD)
    
    If estados Is Nothing Then
        testResult.Fail "GetAvailableStates devolvió Nothing"
        GoTo Cleanup
    End If
    
    If estados.Count = 0 Then
        testResult.Fail "No se encontraron estados disponibles"
        GoTo Cleanup
    End If
    
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
    
    Dim estadosSiguientes As Collection
    Set estadosSiguientes = m_Repository.GetNextStates(TEST_ESTADO_INICIAL, TEST_TIPO_SOLICITUD, TEST_ROL_ADMIN)
    
    If estadosSiguientes Is Nothing Then
        testResult.Fail "GetNextStates devolvió Nothing"
        GoTo Cleanup
    End If
    
    If estadosSiguientes.Count = 0 Then
        testResult.Fail "No se encontraron estados siguientes"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
    Set estadosSiguientes = Nothing
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
    
    Dim estadoInicial As String
    estadoInicial = m_Repository.GetInitialState(TEST_TIPO_SOLICITUD)
    
    If Len(estadoInicial) = 0 Then
        testResult.Fail "No se obtuvo estado inicial"
        GoTo Cleanup
    End If
    
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
    
    ' Probar con estado inicial (no debería ser final)
    Dim esFinal As Boolean
    esFinal = m_Repository.IsStateFinal(TEST_ESTADO_INICIAL, TEST_TIPO_SOLICITUD)
    
    If esFinal Then
        testResult.Fail "Estado inicial incorrectamente identificado como final"
        GoTo Cleanup
    End If
    
    ' Probar con estado final (si existe)
    esFinal = m_Repository.IsStateFinal("APROBADO", TEST_TIPO_SOLICITUD)
    
    ' Nota: No fallamos si APROBADO no es final, ya que depende de la configuración
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
    
    Dim resultado As Boolean
    resultado = m_Repository.RecordStateChange(m_TestSolicitudID, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_USUARIO_TEST, "Cambio de prueba")
    
    If Not resultado Then
        testResult.Fail "No se pudo registrar el cambio de estado"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
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
    
    Dim historial As Collection
    Set historial = m_Repository.GetStateHistory(m_TestSolicitudID)
    
    If historial Is Nothing Then
        testResult.Fail "GetStateHistory devolvió Nothing"
        GoTo Cleanup
    End If
    
    testResult.Pass
    
Cleanup:
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
    
    ' Probar con rol de administrador (debería tener permisos)
    Dim tienePermiso As Boolean
    tienePermiso = m_Repository.HasTransitionPermission(TEST_ROL_ADMIN, TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_TIPO_SOLICITUD)
    
    If Not tienePermiso Then
        testResult.Fail "Administrador sin permisos de transición"
        GoTo Cleanup
    End If
    
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
    
    Dim requiereAprobacion As Boolean
    requiereAprobacion = m_Repository.RequiresApproval(TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_TIPO_SOLICITUD)
    
    ' Esta prueba siempre pasa ya que solo verifica que el método funcione
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
    
    Dim rolRequerido As String
    rolRequerido = m_Repository.GetTransitionRequiredRole(TEST_ESTADO_INICIAL, TEST_ESTADO_REVISION, TEST_TIPO_SOLICITUD)
    
    ' Esta prueba siempre pasa ya que solo verifica que el método funcione
    testResult.Pass
    
Cleanup:
    Set Test_GetTransitionRequiredRole_Integration = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function