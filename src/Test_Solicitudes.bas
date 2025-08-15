Attribute VB_Name = "Test_Solicitudes"
Option Compare Database
Option Explicit

Public Function Test_Solicitudes_RunAll() As String
    Dim resultado As String
    resultado = "=== PRUEBAS DE SOLICITUDES ===" & vbCrLf
    
    Debug.Print "Ejecutando Test_Factory_Crea_PC..."
    On Error GoTo ErrorHandler
    
    Test_Factory_Crea_PC
    resultado = resultado & "✓ Test_Factory_Crea_PC: PASÓ" & vbCrLf
    
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Total de pruebas: 1" & vbCrLf
    resultado = resultado & "Pruebas exitosas: 1" & vbCrLf
    resultado = resultado & "RESULTADO: TODAS LAS PRUEBAS PASARON" & vbCrLf
    
    Test_Solicitudes_RunAll = resultado
    Exit Function
    
ErrorHandler:
    resultado = resultado & "✗ Test_Factory_Crea_PC: FALLÓ - " & Err.Description & vbCrLf
    resultado = resultado & "=== RESUMEN ===" & vbCrLf
    resultado = resultado & "Total de pruebas: 1" & vbCrLf
    resultado = resultado & "Pruebas exitosas: 0" & vbCrLf
    resultado = resultado & "RESULTADO: PRUEBAS FALLARON" & vbCrLf
    Test_Solicitudes_RunAll = resultado
End Function

Private Sub Test_Factory_Crea_PC()
    Dim s As ISolicitud
    Set s = modSolicitudFactory.CreateSolicitud(1)
    Debug.Assert TypeOf s Is CSolicitudPC
End Sub
