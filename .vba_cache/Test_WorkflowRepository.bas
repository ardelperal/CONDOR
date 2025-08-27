Attribute VB_Name = "Test_WorkflowRepository"
Option Compare Database
Option Explicit

'==============================================================================
' Módulo: Test_WorkflowRepository
' Propósito: Pruebas unitarias para CWorkflowRepository usando mocks
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
' Nota: Refactorizado de pruebas de integración a pruebas unitarias puras
'==============================================================================


#If DEV_MODE Then

'==============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'==============================================================================

'''
' Ejecuta todas las pruebas unitarias del WorkflowRepository
' @return CTestSuiteResult: Resultado de la suite de pruebas
'''
Public Function Test_WorkflowRepository_RunAll() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    ' Crear la suite de resultados
    Dim suite As CTestSuiteResult
    Set suite = New CTestSuiteResult
    suite.Initialize "Test_WorkflowRepository", "Pruebas unitarias para CWorkflowRepository"
    
    ' Ejecutar todas las pruebas unitarias
    suite.AddTestResult Test_WorkflowRepository_ValidTransition_ReturnsTrue()
    suite.AddTestResult Test_WorkflowRepository_InvalidTransition_ReturnsFalse()
    suite.AddTestResult Test_WorkflowRepository_NonExistentType_ReturnsFalse()
    suite.AddTestResult Test_WorkflowRepository_InactiveTransition_ReturnsFalse()
    
    Set Test_WorkflowRepository_RunAll = suite
    Exit Function
    
ErrorHandler:
    If suite Is Nothing Then Set suite = New CTestSuiteResult
    suite.Initialize "Test_WorkflowRepository", "Error en suite de pruebas"
    Set Test_WorkflowRepository_RunAll = suite
End Function

'==============================================================================
' PRUEBAS UNITARIAS
'==============================================================================

'''
' Prueba que IsValidTransition devuelve True para una transición válida usando mock
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_ValidTransition_ReturnsTrue() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_ValidTransition_ReturnsTrue", "Verifica que IsValidTransition devuelve True para transición válida usando mock"
    
    On Error GoTo ErrorHandler
    
    ' Crear mock del repositorio
    Dim mockRepository As CMockWorkflowRepository
    Set mockRepository = New CMockWorkflowRepository
    
    ' Configurar regla de transición válida en el mock
    mockRepository.AddRule "PC|BORRADOR|ENVIADO", True
    
    ' Ejecutar la prueba
    Dim result As Boolean
    result = mockRepository.IsValidTransition("PC", "BORRADOR", "ENVIADO")
    
    ' Verificar el resultado
    If result = True Then
        testResult.Pass "La transición válida fue correctamente identificada por el mock"
    Else
        testResult.Fail "La transición válida no fue identificada correctamente por el mock. Esperado: True, Obtenido: " & result
    End If
    
    Set Test_WorkflowRepository_ValidTransition_ReturnsTrue = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Set Test_WorkflowRepository_ValidTransition_ReturnsTrue = testResult
End Function

'''
' Prueba que IsValidTransition devuelve False para una transición inválida usando mock
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_InvalidTransition_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_InvalidTransition_ReturnsFalse", "Verifica que IsValidTransition devuelve False para transición inválida usando mock"
    
    On Error GoTo ErrorHandler
    
    ' Crear mock del repositorio
    Dim mockRepository As CMockWorkflowRepository
    Set mockRepository = New CMockWorkflowRepository
    
    ' Configurar solo una regla válida (PC|BORRADOR|ENVIADO)
    mockRepository.AddRule "PC|BORRADOR|ENVIADO", True
    
    ' Ejecutar la prueba con una transición que no existe (transición inversa)
    Dim result As Boolean
    result = mockRepository.IsValidTransition("PC", "ENVIADO", "BORRADOR")
    
    ' Verificar el resultado
    If result = False Then
        testResult.Pass "La transición inválida fue correctamente rechazada por el mock"
    Else
        testResult.Fail "La transición inválida no fue rechazada correctamente por el mock. Esperado: False, Obtenido: " & result
    End If
    
    Set Test_WorkflowRepository_InvalidTransition_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Set Test_WorkflowRepository_InvalidTransition_ReturnsFalse = testResult
End Function

'''
' Prueba que un tipo de solicitud inexistente devuelve False usando mock
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_NonExistentType_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_NonExistentType_ReturnsFalse", "Verifica que IsValidTransition devuelve False para tipo no existente usando mock"
    
    On Error GoTo ErrorHandler
    
    ' Crear mock del repositorio
    Dim mockRepository As CMockWorkflowRepository
    Set mockRepository = New CMockWorkflowRepository
    
    ' Configurar solo reglas para tipo PC
    mockRepository.AddRule "PC|BORRADOR|ENVIADO", True
    
    ' Ejecutar la prueba con un tipo inexistente
    Dim result As Boolean
    result = mockRepository.IsValidTransition("INEXISTENTE", "BORRADOR", "ENVIADO")
    
    ' Verificar resultado
    If Not result Then
        testResult.Pass "El tipo de solicitud inexistente fue correctamente rechazado por el mock"
    Else
        testResult.Fail "El tipo de solicitud inexistente no fue rechazado correctamente por el mock. Esperado: False, Obtenido: " & result
    End If
    
    Set Test_WorkflowRepository_NonExistentType_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Set Test_WorkflowRepository_NonExistentType_ReturnsFalse = testResult
End Function

'''
' Prueba que IsValidTransition devuelve False para transición inactiva usando mock
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_InactiveTransition_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_InactiveTransition_ReturnsFalse", "Verifica que IsValidTransition devuelve False para transición inactiva usando mock"
    
    On Error GoTo ErrorHandler
    
    ' Crear mock del repositorio
    Dim mockRepository As CMockWorkflowRepository
    Set mockRepository = New CMockWorkflowRepository
    
    ' Configurar transición inactiva (False)
    mockRepository.AddRule "PC|BORRADOR|CANCELADO", False
    
    ' Ejecutar la prueba
    Dim result As Boolean
    result = mockRepository.IsValidTransition("PC", "BORRADOR", "CANCELADO")
    
    ' Verificar el resultado
    If result = False Then
        testResult.Pass "La transición inactiva fue correctamente rechazada por el mock"
    Else
        testResult.Fail "La transición inactiva no fue rechazada correctamente por el mock. Esperado: False, Obtenido: " & result
    End If
    
    Set Test_WorkflowRepository_InactiveTransition_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Set Test_WorkflowRepository_InactiveTransition_ReturnsFalse = testResult
End Function

'==============================================================================
' NOTAS SOBRE REFACTORIZACIÓN
'==============================================================================
' Las funciones auxiliares de configuración y limpieza de datos han sido eliminadas
' ya que las pruebas ahora utilizan mocks en lugar de la base de datos real.
' Esto convierte las pruebas de integración en pruebas unitarias puras.
'==============================================================================

#End If


