Attribute VB_Name = "Test_Debug"
' Módulo de Pruebas de Debug para verificar funcionamiento
' Versión temporal sin compilación condicional

Option Compare Database
Option Explicit

' Función de prueba simple para verificar que el sistema funciona
Public Function TestBasicFunctionality() As String
    Dim resultado As String
    
    resultado = "=== PRUEBA BASICA DE FUNCIONALIDAD ===" & vbCrLf
    
    ' Probar inicialización
    On Error GoTo ErrorHandler
    
    Dim initResult As Boolean
    initResult = modConfig.InitializeEnvironment()
    
    If initResult Then
        resultado = resultado & "✓ InitializeEnvironment: OK" & vbCrLf
    Else
        resultado = resultado & "✗ InitializeEnvironment: FALLO" & vbCrLf
        GoTo ErrorHandler
    End If
    
    ' Probar obtención de rutas
    Dim dbPath As String
    dbPath = modConfig.GetDatabasePath()
    
    If Len(dbPath) > 0 Then
        resultado = resultado & "✓ GetDatabasePath: OK (" & dbPath & ")" & vbCrLf
    Else
        resultado = resultado & "✗ GetDatabasePath: FALLO" & vbCrLf
    End If
    
    ' Probar modo desarrollo
    Dim isDev As Boolean
    isDev = modConfig.IsDevelopmentMode()
    
    resultado = resultado & "✓ IsDevelopmentMode: " & isDev & vbCrLf
    
    resultado = resultado & "=== PRUEBA COMPLETADA EXITOSAMENTE ===" & vbCrLf
    
    TestBasicFunctionality = resultado
    Exit Function
    
ErrorHandler:
    resultado = resultado & "✗ ERROR: " & Err.Number & " - " & Err.Description & vbCrLf
    TestBasicFunctionality = resultado
End Function
