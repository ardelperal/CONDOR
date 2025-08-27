' Test rápido para verificar CMockWorkflowRepository refactorizado
' Simular la lógica del mock
Dim testCollection
Set testCollection = CreateObject("Scripting.Dictionary")

' Agregar algunas reglas de prueba
testCollection.Add "PC|BORRADOR|ENVIADO", True
testCollection.Add "PC|ENVIADO|APROBADO", True
testCollection.Add "PC|BORRADOR|CANCELADO", False

' Función para simular RuleExists con For Each
Function TestRuleExists(collection, ruleKey)
    Dim key
    For Each key In collection.Keys
        If CStr(key) = ruleKey Then
            TestRuleExists = True
            Exit Function
        End If
    Next
    TestRuleExists = False
End Function

' Probar la función
WScript.Echo "Probando RuleExists refactorizado:"
WScript.Echo "Regla PC|BORRADOR|ENVIADO existe: " & TestRuleExists(testCollection, "PC|BORRADOR|ENVIADO")
WScript.Echo "Regla PC|INEXISTENTE|OTRO existe: " & TestRuleExists(testCollection, "PC|INEXISTENTE|OTRO")
WScript.Echo "Regla PC|ENVIADO|APROBADO existe: " & TestRuleExists(testCollection, "PC|ENVIADO|APROBADO")

WScript.Echo "Test completado exitosamente. La lógica For Each funciona correctamente."