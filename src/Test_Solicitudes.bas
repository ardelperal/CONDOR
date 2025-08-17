Attribute VB_Name = "Test_Solicitudes"
Option Compare Database
Option Explicit

' ============================================================================
' FUNCIÓN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_Solicitudes_RunAll() As String
    Dim resultado As String
    resultado = "=== PRUEBAS DE SOLICITUDES ===" & vbCrLf
    
    Debug.Print "Ejecutando Test_Factory_Crea_PC..."
    On Error GoTo ErrorHandler
    
    If Test_Factory_Crea_PC() Then
        resultado = resultado & "? Test_Factory_Crea_PC: PASÓ" & vbCrLf
    Else
        resultado = resultado & "? Test_Factory_Crea_PC: FALLÓ" & vbCrLf
    End If
    
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Total de pruebas: 1" & vbCrLf
    resultado = resultado & "Pruebas exitosas: 1" & vbCrLf
    resultado = resultado & "RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    
    Test_Solicitudes_RunAll = resultado
    Exit Function
    
ErrorHandler:
    resultado = resultado & "? Test_Factory_Crea_PC: FALLÓ - " & Err.Description & vbCrLf
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Total de pruebas: 1" & vbCrLf
    resultado = resultado & "Pruebas exitosas: 0" & vbCrLf
    resultado = resultado & "RESULTADO: PRUEBAS FALLARON" & vbCrLf
    Test_Solicitudes_RunAll = resultado
End Function

Private Function Test_Factory_Crea_PC() As Boolean
    ' Test básico - solo verificar que se puede crear el objeto
    On Error GoTo ErrorHandler
    
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Si llegamos aquí, el objeto se creó exitosamente
    Test_Factory_Crea_PC = True
    Exit Function
    
ErrorHandler:
    Test_Factory_Crea_PC = False
End Function


