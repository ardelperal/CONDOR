Attribute VB_Name = "modAssert"
Option Compare Database
Option Explicit

' Modulo: modAssert
' Proposito: Funciones de asercion para las pruebas.

' ============================================================================
' FUNCIONES DE ASERCIÓN PARA VALORES BOOLEANOS
' ============================================================================

' Función: AssertTrue
' Propósito: Verifica que una condición sea verdadera
' Parámetros:
'   - condition: La condición booleana a verificar
'   - message: Mensaje descriptivo para mostrar si la aserción falla
' Error: Lanza vbObjectError + 510 si la condición es falsa
Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then
        Err.Raise vbObjectError + 510, "modAssert.AssertTrue", "Assertion Failed: " & message
    End If
End Sub

' Función: AssertFalse
' Propósito: Verifica que una condición sea falsa
' Parámetros:
'   - condition: La condición booleana a verificar
'   - message: Mensaje descriptivo para mostrar si la aserción falla
' Error: Lanza vbObjectError + 511 si la condición es verdadera
Public Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    If condition Then
        Err.Raise vbObjectError + 511, "modAssert.AssertFalse", "Assertion Failed: " & message
    End If
End Sub

' Función: IsTrue (Mantenida para compatibilidad hacia atrás)
' Propósito: Función legacy que solo imprime en Debug
Public Sub IsTrue(value As Boolean, message As String)
    If Not value Then
        Debug.Print "ASSERT FAILED: " & message
    End If
End Sub

' ============================================================================
' FUNCIONES DE ASERCIÓN PARA OBJETOS
' ============================================================================

' Función: AssertNotNull
' Propósito: Verifica que un objeto no sea Nothing
' Parámetros:
'   - obj: El objeto a verificar
'   - message: Mensaje descriptivo para mostrar si la aserción falla
' Error: Lanza vbObjectError + 513 si el objeto es Nothing
Public Sub AssertNotNull(ByVal obj As Object, ByVal message As String)
    If obj Is Nothing Then
        Err.Raise vbObjectError + 513, "modAssert.AssertNotNull", "Assertion Failed: " & message
    End If
End Sub

' Función: AssertIsNull
' Propósito: Verifica que un objeto sea Nothing
' Parámetros:
'   - obj: El objeto a verificar
'   - message: Mensaje descriptivo para mostrar si la aserción falla
' Error: Lanza vbObjectError + 514 si el objeto no es Nothing
Public Sub AssertIsNull(ByVal obj As Object, ByVal message As String)
    If Not obj Is Nothing Then
        Err.Raise vbObjectError + 514, "modAssert.AssertIsNull", "Assertion Failed: " & message
    End If
End Sub

' ============================================================================
' FUNCIONES DE ASERCIÓN PARA IGUALDAD
' ============================================================================

' Función: AssertEquals
' Propósito: Verifica que dos valores sean iguales
' Parámetros:
'   - expected: El valor esperado
'   - actual: El valor actual a comparar
'   - message: Mensaje descriptivo para mostrar si la aserción falla
' Error: Lanza vbObjectError + 512 si los valores no son iguales
Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 512, "modAssert.AssertEquals", "Assertion Failed: " & message & " (Expected: '" & CStr(expected) & "', Actual: '" & CStr(actual) & "')"
    End If
End Sub

' ============================================================================
' FUNCIONES DE ASERCIÓN GENERALES
' ============================================================================

' Función: Fail
' Propósito: Fuerza un fallo de aserción incondicional
' Parámetros:
'   - message: Mensaje descriptivo del fallo
' Error: Siempre lanza vbObjectError + 515
Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 515, "modAssert.Fail", "Test Failed: " & message
End Sub


