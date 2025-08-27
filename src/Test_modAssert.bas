Attribute VB_Name = "Test_modAssert"
Option Compare Database
Option Explicit


#If DEV_MODE Then

' ============================================================================
' MÓDULO DE META-TESTING PARA modAssert
' ============================================================================
' Este módulo contiene pruebas unitarias para el propio framework de aserciones.
' Verifica que cada función de aserción funcione correctamente tanto en casos
' de éxito como de fallo.

' Función principal que ejecuta todas las pruebas del módulo
Public Function Test_modAssert_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_modAssert"
    
    ' Ejecutar todas las pruebas de meta-testing
    suiteResult.AddTestResult Test_AssertTrue_WithTrueCondition_Passes()
    suiteResult.AddTestResult Test_AssertTrue_WithFalseCondition_Fails()
    suiteResult.AddTestResult Test_AssertFalse_WithFalseCondition_Passes()
    suiteResult.AddTestResult Test_AssertFalse_WithTrueCondition_Fails()
    suiteResult.AddTestResult Test_AssertEquals_WithEqualValues_Passes()
    suiteResult.AddTestResult Test_AssertEquals_WithDifferentValues_Fails()
    suiteResult.AddTestResult Test_AssertNotNull_WithValidObject_Passes()
    suiteResult.AddTestResult Test_AssertNotNull_WithNothingObject_Fails()
    suiteResult.AddTestResult Test_AssertIsNull_WithNothingObject_Passes()
    suiteResult.AddTestResult Test_AssertIsNull_WithValidObject_Fails()
    suiteResult.AddTestResult Test_Fail_AlwaysFails()
    
    Set Test_modAssert_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS PARA AssertTrue
' ============================================================================

' Prueba que AssertTrue no falla cuando se le pasa True
Private Function Test_AssertTrue_WithTrueCondition_Passes() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertTrue_WithTrueCondition_Passes"
    
    On Error GoTo ErrorHandler
    
    ' Act & Assert - No debe lanzar error
    modAssert.AssertTrue True, "Esta aserción debe pasar"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "AssertTrue falló inesperadamente con condición True: " & Err.Description
    
Cleanup:
    Set Test_AssertTrue_WithTrueCondition_Passes = testResult
End Function

' Prueba que AssertTrue falla cuando se le pasa False
Private Function Test_AssertTrue_WithFalseCondition_Fails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertTrue_WithFalseCondition_Fails"
    
    On Error GoTo ErrorHandler
    
    ' Act - Debe lanzar error
    modAssert.AssertTrue False, "Esta aserción debe fallar"
    
    ' Si llegamos aquí, la aserción no falló como debería
    testResult.Fail "AssertTrue debería haber fallado con condición False"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 510 Then
        testResult.Pass
    Else
        testResult.Fail "AssertTrue falló con error incorrecto. Esperado: " & (vbObjectError + 510) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set Test_AssertTrue_WithFalseCondition_Fails = testResult
End Function

' ============================================================================
' PRUEBAS PARA AssertFalse
' ============================================================================

' Prueba que AssertFalse no falla cuando se le pasa False
Private Function Test_AssertFalse_WithFalseCondition_Passes() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertFalse_WithFalseCondition_Passes"
    
    On Error GoTo ErrorHandler
    
    ' Act & Assert - No debe lanzar error
    modAssert.AssertFalse False, "Esta aserción debe pasar"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "AssertFalse falló inesperadamente con condición False: " & Err.Description
    
Cleanup:
    Set Test_AssertFalse_WithFalseCondition_Passes = testResult
End Function

' Prueba que AssertFalse falla cuando se le pasa True
Private Function Test_AssertFalse_WithTrueCondition_Fails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertFalse_WithTrueCondition_Fails"
    
    On Error GoTo ErrorHandler
    
    ' Act - Debe lanzar error
    modAssert.AssertFalse True, "Esta aserción debe fallar"
    
    ' Si llegamos aquí, la aserción no falló como debería
    testResult.Fail "AssertFalse debería haber fallado con condición True"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 511 Then
        testResult.Pass
    Else
        testResult.Fail "AssertFalse falló con error incorrecto. Esperado: " & (vbObjectError + 511) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set Test_AssertFalse_WithTrueCondition_Fails = testResult
End Function

' ============================================================================
' PRUEBAS PARA AssertEquals
' ============================================================================

' Prueba que AssertEquals no falla cuando los valores son iguales
Private Function Test_AssertEquals_WithEqualValues_Passes() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertEquals_WithEqualValues_Passes"
    
    On Error GoTo ErrorHandler
    
    ' Act & Assert - No debe lanzar error
    modAssert.AssertEquals "test", "test", "Valores iguales deben pasar"
    modAssert.AssertEquals 42, 42, "Números iguales deben pasar"
    modAssert.AssertEquals True, True, "Booleanos iguales deben pasar"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "AssertEquals falló inesperadamente con valores iguales: " & Err.Description
    
Cleanup:
    Set Test_AssertEquals_WithEqualValues_Passes = testResult
End Function

' Prueba que AssertEquals falla cuando los valores son diferentes
Private Function Test_AssertEquals_WithDifferentValues_Fails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertEquals_WithDifferentValues_Fails"
    
    On Error GoTo ErrorHandler
    
    ' Act - Debe lanzar error
    modAssert.AssertEquals "expected", "actual", "Esta aserción debe fallar"
    
    ' Si llegamos aquí, la aserción no falló como debería
    testResult.Fail "AssertEquals debería haber fallado con valores diferentes"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 512 Then
        testResult.Pass
    Else
        testResult.Fail "AssertEquals falló con error incorrecto. Esperado: " & (vbObjectError + 512) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set Test_AssertEquals_WithDifferentValues_Fails = testResult
End Function

' ============================================================================
' PRUEBAS PARA AssertNotNull
' ============================================================================

' Prueba que AssertNotNull no falla cuando el objeto no es Nothing
Private Function Test_AssertNotNull_WithValidObject_Passes() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertNotNull_WithValidObject_Passes"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim obj As New Collection
    
    ' Act & Assert - No debe lanzar error
    modAssert.AssertNotNull obj, "Objeto válido debe pasar"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "AssertNotNull falló inesperadamente con objeto válido: " & Err.Description
    
Cleanup:
    Set obj = Nothing
    Set Test_AssertNotNull_WithValidObject_Passes = testResult
End Function

' Prueba que AssertNotNull falla cuando el objeto es Nothing
Private Function Test_AssertNotNull_WithNothingObject_Fails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertNotNull_WithNothingObject_Fails"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim obj As Object
    Set obj = Nothing
    
    ' Act - Debe lanzar error
    modAssert.AssertNotNull obj, "Esta aserción debe fallar"
    
    ' Si llegamos aquí, la aserción no falló como debería
    testResult.Fail "AssertNotNull debería haber fallado con objeto Nothing"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 513 Then
        testResult.Pass
    Else
        testResult.Fail "AssertNotNull falló con error incorrecto. Esperado: " & (vbObjectError + 513) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set Test_AssertNotNull_WithNothingObject_Fails = testResult
End Function

' ============================================================================
' PRUEBAS PARA AssertIsNull
' ============================================================================

' Prueba que AssertIsNull no falla cuando el objeto es Nothing
Private Function Test_AssertIsNull_WithNothingObject_Passes() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertIsNull_WithNothingObject_Passes"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim obj As Object
    Set obj = Nothing
    
    ' Act & Assert - No debe lanzar error
    modAssert.AssertIsNull obj, "Objeto Nothing debe pasar"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "AssertIsNull falló inesperadamente con objeto Nothing: " & Err.Description
    
Cleanup:
    Set Test_AssertIsNull_WithNothingObject_Passes = testResult
End Function

' Prueba que AssertIsNull falla cuando el objeto no es Nothing
Private Function Test_AssertIsNull_WithValidObject_Fails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_AssertIsNull_WithValidObject_Fails"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim obj As New Collection
    
    ' Act - Debe lanzar error
    modAssert.AssertIsNull obj, "Esta aserción debe fallar"
    
    ' Si llegamos aquí, la aserción no falló como debería
    testResult.Fail "AssertIsNull debería haber fallado con objeto válido"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 514 Then
        testResult.Pass
    Else
        testResult.Fail "AssertIsNull falló con error incorrecto. Esperado: " & (vbObjectError + 514) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set obj = Nothing
    Set Test_AssertIsNull_WithValidObject_Fails = testResult
End Function

' ============================================================================
' PRUEBAS PARA Fail
' ============================================================================

' Prueba que Fail siempre falla
Private Function Test_Fail_AlwaysFails() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Fail_AlwaysFails"
    
    On Error GoTo ErrorHandler
    
    ' Act - Debe lanzar error
    modAssert.Fail "Esta función siempre debe fallar"
    
    ' Si llegamos aquí, Fail no falló como debería
    testResult.Fail "Fail debería haber fallado incondicionalmente"
    GoTo Cleanup
    
ErrorHandler:
    ' Verificar que el error sea el esperado
    If Err.Number = vbObjectError + 515 Then
        testResult.Pass
    Else
        testResult.Fail "Fail falló con error incorrecto. Esperado: " & (vbObjectError + 515) & ", Actual: " & Err.Number
    End If
    
Cleanup:
    Set Test_Fail_AlwaysFails = testResult
End Function

#End If

