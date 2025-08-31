Attribute VB_Name = "TestModAssert"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE META-TESTING PARA modAssert
' ============================================================================
' Este módulo contiene pruebas unitarias para el propio framework de aserciones.
' Verifica que cada función de aserción funcione correctamente tanto en casos
' de éxito como de fallo.

' Función principal que ejecuta todas las pruebas del módulo
Public Function TestModAssertRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestModAssert")
    
    ' Ejecutar todas las pruebas de meta-testing
    Call suiteResult.AddTestResult(TestAssertTrueWithTrueConditionPasses())
    Call suiteResult.AddTestResult(TestAssertTrueWithFalseConditionFails())
    Call suiteResult.AddTestResult(TestAssertFalseWithFalseConditionPasses())
    Call suiteResult.AddTestResult(TestAssertFalseWithTrueConditionFails())
    Call suiteResult.AddTestResult(TestAssertEqualsWithEqualValuesPasses())
    Call suiteResult.AddTestResult(TestAssertEqualsWithDifferentValuesFails())
    Call suiteResult.AddTestResult(TestAssertNotNullWithValidObjectPasses())
    Call suiteResult.AddTestResult(TestAssertNotNullWithNothingObjectFails())
    Call suiteResult.AddTestResult(TestAssertIsNullWithNothingObjectPasses())
    Call suiteResult.AddTestResult(TestAssertIsNullWithValidObjectFails())
    Call suiteResult.AddTestResult(TestFailAlwaysFails())
    
    Set TestModAssertRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS PARA AssertTrue
' ============================================================================

Private Function TestAssertTrueWithTrueConditionPasses() As CTestResult
    Set TestAssertTrueWithTrueConditionPasses = New CTestResult
    TestAssertTrueWithTrueConditionPasses.Initialize "AssertTrue con condición Verdadera debe pasar"
    On Error GoTo TestFail
    
    ModAssert.AssertTrue True, "Esta aserción debe pasar"
    
    TestAssertTrueWithTrueConditionPasses.Pass
    Exit Function
TestFail:
    TestAssertTrueWithTrueConditionPasses.Fail "AssertTrue falló inesperadamente con condición True: " & Err.Description
End Function

Private Function TestAssertTrueWithFalseConditionFails() As CTestResult
    Set TestAssertTrueWithFalseConditionFails = New CTestResult
    TestAssertTrueWithFalseConditionFails.Initialize "AssertTrue con condición Falsa debe fallar"
    On Error GoTo TestFail
    
    ModAssert.AssertTrue False, "Esta aserción debe fallar"
    
    TestAssertTrueWithFalseConditionFails.Fail "AssertTrue debería haber fallado con condición False"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 510 Then
        TestAssertTrueWithFalseConditionFails.Pass
    Else
        TestAssertTrueWithFalseConditionFails.Fail "AssertTrue falló con error incorrecto. Esperado: " & (vbObjectError + 510) & ", Actual: " & Err.Number
    End If
End Function

' ============================================================================
' PRUEBAS PARA AssertFalse
' ============================================================================

Private Function TestAssertFalseWithFalseConditionPasses() As CTestResult
    Set TestAssertFalseWithFalseConditionPasses = New CTestResult
    TestAssertFalseWithFalseConditionPasses.Initialize "AssertFalse con condición Falsa debe pasar"
    On Error GoTo TestFail
    
    ModAssert.AssertFalse False, "Esta aserción debe pasar"
    
    TestAssertFalseWithFalseConditionPasses.Pass
    Exit Function
TestFail:
    TestAssertFalseWithFalseConditionPasses.Fail "AssertFalse falló inesperadamente con condición False: " & Err.Description
End Function

Private Function TestAssertFalseWithTrueConditionFails() As CTestResult
    Set TestAssertFalseWithTrueConditionFails = New CTestResult
    TestAssertFalseWithTrueConditionFails.Initialize "AssertFalse con condición Verdadera debe fallar"
    On Error GoTo TestFail
    
    ModAssert.AssertFalse True, "Esta aserción debe fallar"
    
    TestAssertFalseWithTrueConditionFails.Fail "AssertFalse debería haber fallado con condición True"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 511 Then
        TestAssertFalseWithTrueConditionFails.Pass
    Else
        TestAssertFalseWithTrueConditionFails.Fail "AssertFalse falló con error incorrecto. Esperado: " & (vbObjectError + 511) & ", Actual: " & Err.Number
    End If
End Function

' ============================================================================
' PRUEBAS PARA AssertEquals
' ============================================================================

Private Function TestAssertEqualsWithEqualValuesPasses() As CTestResult
    Set TestAssertEqualsWithEqualValuesPasses = New CTestResult
    TestAssertEqualsWithEqualValuesPasses.Initialize "AssertEquals con valores iguales debe pasar"
    On Error GoTo TestFail
    
    ModAssert.AssertEquals "test", "test", "Valores iguales deben pasar"
    ModAssert.AssertEquals 42, 42, "Números iguales deben pasar"
    ModAssert.AssertEquals True, True, "Booleanos iguales deben pasar"
    
    TestAssertEqualsWithEqualValuesPasses.Pass
    Exit Function
TestFail:
    TestAssertEqualsWithEqualValuesPasses.Fail "AssertEquals falló inesperadamente con valores iguales: " & Err.Description
End Function

Private Function TestAssertEqualsWithDifferentValuesFails() As CTestResult
    Set TestAssertEqualsWithDifferentValuesFails = New CTestResult
    TestAssertEqualsWithDifferentValuesFails.Initialize "AssertEquals con valores diferentes debe fallar"
    On Error GoTo TestFail
    
    ModAssert.AssertEquals "expected", "actual", "Esta aserción debe fallar"
    
    TestAssertEqualsWithDifferentValuesFails.Fail "AssertEquals debería haber fallado con valores diferentes"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 512 Then
        TestAssertEqualsWithDifferentValuesFails.Pass
    Else
        TestAssertEqualsWithDifferentValuesFails.Fail "AssertEquals falló con error incorrecto. Esperado: " & (vbObjectError + 512) & ", Actual: " & Err.Number
    End If
End Function

' ============================================================================
' PRUEBAS PARA AssertNotNull
' ============================================================================

Private Function TestAssertNotNullWithValidObjectPasses() As CTestResult
    Set TestAssertNotNullWithValidObjectPasses = New CTestResult
    TestAssertNotNullWithValidObjectPasses.Initialize "AssertNotNull con un objeto válido debe pasar"
    On Error GoTo TestFail
    
    Dim obj As New Collection
    ModAssert.AssertNotNull obj, "Objeto válido debe pasar"
    
    TestAssertNotNullWithValidObjectPasses.Pass
    Exit Function
TestFail:
    TestAssertNotNullWithValidObjectPasses.Fail "AssertNotNull falló inesperadamente con objeto válido: " & Err.Description
End Function

Private Function TestAssertNotNullWithNothingObjectFails() As CTestResult
    Set TestAssertNotNullWithNothingObjectFails = New CTestResult
    TestAssertNotNullWithNothingObjectFails.Initialize "AssertNotNull con un objeto Nothing debe fallar"
    On Error GoTo TestFail
    
    Dim obj As Object
    Set obj = Nothing
    ModAssert.AssertNotNull obj, "Esta aserción debe fallar"
    
    TestAssertNotNullWithNothingObjectFails.Fail "AssertNotNull debería haber fallado con objeto Nothing"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 513 Then
        TestAssertNotNullWithNothingObjectFails.Pass
    Else
        TestAssertNotNullWithNothingObjectFails.Fail "AssertNotNull falló con error incorrecto. Esperado: " & (vbObjectError + 513) & ", Actual: " & Err.Number
    End If
End Function

' ============================================================================
' PRUEBAS PARA AssertIsNull
' ============================================================================

Private Function TestAssertIsNullWithNothingObjectPasses() As CTestResult
    Set TestAssertIsNullWithNothingObjectPasses = New CTestResult
    TestAssertIsNullWithNothingObjectPasses.Initialize "AssertIsNull con un objeto Nothing debe pasar"
    On Error GoTo TestFail
    
    Dim obj As Object
    Set obj = Nothing
    ModAssert.AssertIsNull obj, "Objeto Nothing debe pasar"
    
    TestAssertIsNullWithNothingObjectPasses.Pass
    Exit Function
TestFail:
    TestAssertIsNullWithNothingObjectPasses.Fail "AssertIsNull falló inesperadamente con objeto Nothing: " & Err.Description
End Function

Private Function TestAssertIsNullWithValidObjectFails() As CTestResult
    Set TestAssertIsNullWithValidObjectFails = New CTestResult
    TestAssertIsNullWithValidObjectFails.Initialize "AssertIsNull con un objeto válido debe fallar"
    On Error GoTo TestFail
    
    Dim obj As New Collection
    ModAssert.AssertIsNull obj, "Esta aserción debe fallar"
    
    TestAssertIsNullWithValidObjectFails.Fail "AssertIsNull debería haber fallado con objeto válido"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 514 Then
        TestAssertIsNullWithValidObjectFails.Pass
    Else
        TestAssertIsNullWithValidObjectFails.Fail "AssertIsNull falló con error incorrecto. Esperado: " & (vbObjectError + 514) & ", Actual: " & Err.Number
    End If
End Function

' ============================================================================
' PRUEBAS PARA Fail
' ============================================================================

Private Function TestFailAlwaysFails() As CTestResult
    Set TestFailAlwaysFails = New CTestResult
    TestFailAlwaysFails.Initialize "Fail siempre debe fallar"
    On Error GoTo TestFail
    
    Call ModAssert.Fail("Esta función siempre debe fallar")
    
    TestFailAlwaysFails.Fail "Fail debería haber fallado incondicionalmente"
    Exit Function
TestFail:
    If Err.Number = vbObjectError + 515 Then
        TestFailAlwaysFails.Pass
    Else
        TestFailAlwaysFails.Fail "Fail falló con error incorrecto. Esperado: " & (vbObjectError + 515) & ", Actual: " & Err.Number
    End If
End Function