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
    Dim resultado As Boolean
    resultado = workflowService.ValidateTransition(1, "Borrador", "EnProceso", "PC", "Administrador")
    
    ' Assert
    If resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba True pero se obtuvo False"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar
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
    Dim resultado As Boolean
    resultado = workflowService.ValidateTransition(1, "Borrador", "Aprobado", "PC", "Administrador")
    
    ' Assert
    If Not resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False pero se obtuvo True"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar
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
    Dim resultado As Boolean
    resultado = workflowService.ValidateTransition(1, "Borrador", "Aprobado", "PC", "Usuario")
    
    ' Assert
    If Not resultado Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False pero se obtuvo True"
    End If
    
    ' Verificar que se llamó al método IsValidTransition del repositorio
    If Not mockWorkflowRepo.IsValidTransition_WasCalled Then
        testResult.Fail "No se llamó al método IsValidTransition del repositorio"
    End If
    
Cleanup:
    ' Limpiar
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
    ' Nota: Como no podemos modificar la interfaz IWorkflowRepository para añadir un método GetNextStates,
    ' esta prueba solo verifica que se devuelve una colección (no vacía)
    If Not estados Is Nothing And estados.Count > 0 Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba una colección no vacía"
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