Attribute VB_Name = "Test_WorkflowService"
Option Compare Database
Option Explicit


'******************************************************************************
' MÓDULO: Test_WorkflowService
' DESCRIPCIÓN: Pruebas unitarias para CWorkflowService
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTAS: Lección 10 - El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable
'******************************************************************************

'******************************************************************************
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'******************************************************************************

' Ejecuta todas las pruebas del WorkflowService
Public Function Test_WorkflowService_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "Test_WorkflowService"
    
    ' Ejecutar todas las pruebas
    suite.AddTest Test_ValidateTransition_Valid
    suite.AddTest Test_ValidateTransition_Invalid
    suite.AddTest Test_ValidateTransition_NoPermissions
    suite.AddTest Test_GetNextStates
    suite.AddTest Test_GetAvailableStates
    suite.AddTest Test_GetInitialState
    suite.AddTest Test_IsStateFinal
    suite.AddTest Test_RecordStateChange
    suite.AddTest Test_GetStateHistory
    suite.AddTest Test_HasTransitionPermission
    suite.AddTest Test_RequiresApproval
    
    Set Test_WorkflowService_RunAll = suite
End Function

'******************************************************************************
' PRUEBAS UNITARIAS
'******************************************************************************

' Test que verifica que una transición válida retorna True
Private Function Test_ValidateTransition_Valid() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ValidateTransition con transición válida retorna True"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Configurar el mock para que IsValidTransition retorne True
    mockWorkflowRepo.AddRule "PC", "Borrador", "EnProceso", True
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim Resultado As Boolean
    Resultado = workflowService.ValidateTransition(1, "Borrador", "EnProceso", "PC", "Administrador")
    
    ' Assert
    If Resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba True pero se obtuvo False"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar mocks
    mockConfig.Reset
    mockLogger.Reset
    mockWorkflowRepo.Reset
    
    ' Limpiar referencias
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_ValidateTransition_Valid = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que una transición inválida retorna False
Private Function Test_ValidateTransition_Invalid() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ValidateTransition con transición inválida retorna False"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Configurar el mock para que IsValidTransition retorne False
    mockWorkflowRepo.AddRule "PC", "Borrador", "Aprobado", False
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim Resultado As Boolean
    Resultado = workflowService.ValidateTransition(1, "Borrador", "Aprobado", "PC", "Administrador")
    
    ' Assert
    If Not Resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False pero se obtuvo True"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar mocks
    mockConfig.Reset
    mockLogger.Reset
    mockWorkflowRepo.Reset
    
    ' Limpiar referencias
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_ValidateTransition_Invalid = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que una transición sin los permisos correctos retorna False
Private Function Test_ValidateTransition_NoPermissions() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "ValidateTransition sin permisos retorna False"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Configurar el mock para que IsValidTransition retorne True (la transición es válida)
    mockWorkflowRepo.AddRule "PC", "Borrador", "Aprobado", True
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act - Usamos un rol sin permisos (Usuario)
    Dim Resultado As Boolean
    Resultado = workflowService.ValidateTransition(1, "Borrador", "Aprobado", "PC", "Usuario")
    
    ' Assert
    If Not Resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False pero se obtuvo True"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar mocks
    mockConfig.Reset
    mockLogger.Reset
    mockWorkflowRepo.Reset
    
    ' Limpiar referencias
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_ValidateTransition_NoPermissions = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que GetNextStates devuelve la colección correcta
Private Function Test_GetNextStates() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetNextStates devuelve colección correcta"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim estados As Collection
    Set estados = workflowService.GetNextStates("Borrador", "PC", "Administrador")
    
    ' Assert
    If Not estados Is Nothing Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba una colección válida"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_GetNextStates = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que GetAvailableStates devuelve la colección correcta
Private Function Test_GetAvailableStates() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetAvailableStates devuelve colección correcta"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim estados As Collection
    Set estados = workflowService.GetAvailableStates("PC")
    
    ' Assert
    If Not estados Is Nothing Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba una colección válida"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_GetAvailableStates = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que GetInitialState devuelve el estado inicial correcto
Private Function Test_GetInitialState() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetInitialState devuelve estado inicial correcto"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim estadoInicial As String
    estadoInicial = workflowService.GetInitialState("PC")
    
    ' Assert
    If Len(estadoInicial) > 0 Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba un estado inicial válido"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_GetInitialState = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que IsStateFinal funciona correctamente
Private Function Test_IsStateFinal() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "IsStateFinal funciona correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim esFinal As Boolean
    esFinal = workflowService.IsStateFinal("Aprobado", "PC")
    
    ' Assert - El mock devuelve True para estados finales
    If esFinal Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba que el estado fuera final"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_IsStateFinal = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que RecordStateChange funciona correctamente
Private Function Test_RecordStateChange() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "RecordStateChange funciona correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act - No debería generar error
    workflowService.RecordStateChange 1, "Borrador", "EnProceso", "TestUser", "Comentario de prueba"
    
    ' Assert - Si llegamos aquí sin error, la prueba pasa
    testResult.Pass
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_RecordStateChange = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que GetStateHistory devuelve el historial correcto
Private Function Test_GetStateHistory() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "GetStateHistory devuelve historial correcto"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim historial As Collection
    Set historial = workflowService.GetStateHistory(1)
    
    ' Assert
    If Not historial Is Nothing Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba una colección válida"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_GetStateHistory = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que HasTransitionPermission funciona correctamente
Private Function Test_HasTransitionPermission() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "HasTransitionPermission funciona correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim tienePermiso As Boolean
    tienePermiso = workflowService.HasTransitionPermission("Borrador", "EnProceso", "PC", "Administrador")
    
    ' Assert - El mock devuelve True para permisos
    If tienePermiso Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba que tuviera permiso"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_HasTransitionPermission = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

' Test que verifica que RequiresApproval funciona correctamente
Private Function Test_RequiresApproval() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "RequiresApproval funciona correctamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim workflowService As New CWorkflowService
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockWorkflowRepo As New CMockWorkflowRepository
    
    ' Inicializar el servicio con los mocks
    workflowService.Initialize mockConfig, mockLogger, mockWorkflowRepo
    
    ' Act
    Dim requiereAprobacion As Boolean
    requiereAprobacion = workflowService.RequiresApproval("Borrador", "Aprobado", "PC")
    
    ' Assert - El mock devuelve False para aprobaciones
    If Not requiereAprobacion Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba que no requiriera aprobación"
    End If
    
Cleanup:
    ' Limpiar
    Set workflowService = Nothing
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockWorkflowRepo = Nothing
    
    Set Test_RequiresApproval = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    Resume Cleanup
End Function

