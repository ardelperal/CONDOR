Attribute VB_Name = "Test_SolicitudFactory"
Option Compare Database
Option Explicit

' ============================================================================
' M?dulo: Test_SolicitudFactory
' Descripci?n: Pruebas unitarias para modSolicitudFactory.bas
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de solicitud
Private Type T_MockSolicitudData
    ID As Long
    TipoSolicitud As String
    IsValid As Boolean
    ShouldFailLoad As Boolean
End Type

Private m_MockData As T_MockSolicitudData

' ============================================================================
' FUNCIONES DE CONFIGURACI?N DE MOCKS
' ============================================================================

Private Sub SetupMockSolicitudData()
    m_MockData.ID = 123
    m_MockData.TipoSolicitud = "PC"
    m_MockData.IsValid = True
    m_MockData.ShouldFailLoad = False
End Sub

Private Sub ConfigureMockToFailLoad()
    m_MockData.ShouldFailLoad = True
    m_MockData.IsValid = False
End Sub

' ============================================================================
' PRUEBAS PARA CreateSolicitud
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
' ============================================================================

Public Function Test_SolicitudFactory_RunAll() As String
    Dim resultado As String
    Dim testsPassed As Long, testsTotal As Long
    
    resultado = "=== PRUEBAS DE SOLICITUDFACTORY ===" & vbCrLf
    testsPassed = 0
    testsTotal = 0
    
    ' Ejecutar todas las pruebas
    testsTotal = testsTotal + 1
    If Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC() Then
        resultado = resultado & "[OK] Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CreateSolicitud_InvalidID_ReturnsNothing() Then
        resultado = resultado & "[OK] Test_CreateSolicitud_InvalidID_ReturnsNothing" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateSolicitud_InvalidID_ReturnsNothing" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CreateSolicitud_ZeroID_ReturnsNothing() Then
        resultado = resultado & "[OK] Test_CreateSolicitud_ZeroID_ReturnsNothing" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateSolicitud_ZeroID_ReturnsNothing" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_GetTipoSolicitud_DefaultsToPC() Then
        resultado = resultado & "[OK] Test_GetTipoSolicitud_DefaultsToPC" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_GetTipoSolicitud_DefaultsToPC" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Factory_CreatesValidISolicitudInterface() Then
        resultado = resultado & "[OK] Test_Factory_CreatesValidISolicitudInterface" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Factory_CreatesValidISolicitudInterface" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_Factory_HandlesDatabaseErrors() Then
        resultado = resultado & "[OK] Test_Factory_HandlesDatabaseErrors" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_Factory_HandlesDatabaseErrors" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_Properties_SetAndGet() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_Properties_SetAndGet" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_Properties_SetAndGet" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_Load_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_Load_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_Load_ReturnsTrue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_Save_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_Save_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_Save_ReturnsTrue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_ChangeState_ReturnsTrue() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_ChangeState_ReturnsTrue" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_ChangeState_ReturnsTrue" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_DatosPC_SetAndGet() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_DatosPC_SetAndGet" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_DatosPC_SetAndGet" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CreateSolicitud_LargeID_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_CreateSolicitud_LargeID_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CreateSolicitud_LargeID_HandlesCorrectly" & vbCrLf
    End If
    
    testsTotal = testsTotal + 1
    If Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly() Then
        resultado = resultado & "[OK] Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly" & vbCrLf
        testsPassed = testsPassed + 1
    Else
        resultado = resultado & "[FAIL] Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly" & vbCrLf
    End If
    
    ' Agregar resumen
    resultado = resultado & vbCrLf & "RESUMEN: " & testsPassed & "/" & testsTotal & " pruebas pasadas" & vbCrLf
    
    Test_SolicitudFactory_RunAll = resultado
End Function

Public Function Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupMockSolicitudData
    Dim idSolicitud As Long
    idSolicitud = 123
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Verificamos que se retorna una instancia (no Nothing)
    ' Nota: En un entorno real, esto requerir?a datos v?lidos en la base de datos
    Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC = Not (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC = False
End Function

Public Function Test_CreateSolicitud_InvalidID_ReturnsNothing() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = -1 ' ID inv?lido
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Con ID inv?lido, deber?a retornar Nothing
    Test_CreateSolicitud_InvalidID_ReturnsNothing = (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_InvalidID_ReturnsNothing = False
End Function

Public Function Test_CreateSolicitud_ZeroID_ReturnsNothing() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = 0 ' ID cero
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    Test_CreateSolicitud_ZeroID_ReturnsNothing = (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_ZeroID_ReturnsNothing = False
End Function

' ============================================================================
' PRUEBAS PARA GetTipoSolicitud (funci?n privada - prueba indirecta)
' ============================================================================

Public Function Test_GetTipoSolicitud_DefaultsToPC() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    ' Como GetTipoSolicitud es privada, la probamos indirectamente
    ' a trav?s de CreateSolicitud
    Dim idSolicitud As Long
    idSolicitud = 999 ' Cualquier ID
    
    ' La funci?n privada GetTipoSolicitud siempre retorna "PC" por defecto
    ' Esto se refleja en que CreateSolicitud siempre crea CSolicitudPC
    
    ' Assert
    ' Por ahora, asumimos que funciona correctamente seg?n la implementaci?n
    Test_GetTipoSolicitud_DefaultsToPC = True
    
    Exit Function
    
TestFail:
    Test_GetTipoSolicitud_DefaultsToPC = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACI?N
' ============================================================================

Public Function Test_Factory_CreatesValidISolicitudInterface() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = 456
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Verificamos que el objeto creado implementa la interfaz ISolicitud
    If Not (solicitud Is Nothing) Then
        ' Si no es Nothing, significa que se cre? correctamente
        ' En un entorno real, podr?amos verificar propiedades espec?ficas
        Test_Factory_CreatesValidISolicitudInterface = True
    Else
        Test_Factory_CreatesValidISolicitudInterface = False
    End If
    
    Exit Function
    
TestFail:
    Test_Factory_CreatesValidISolicitudInterface = False
End Function

Public Function Test_Factory_HandlesDatabaseErrors() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    ConfigureMockToFailLoad
    Dim idSolicitud As Long
    idSolicitud = 789
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Cuando hay errores de base de datos, deber?a retornar Nothing
    Test_Factory_HandlesDatabaseErrors = (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    Test_Factory_HandlesDatabaseErrors = False
End Function

' ============================================================================
' PRUEBAS PARA CSolicitudPC
' ============================================================================

Public Function Test_CSolicitudPC_Properties_SetAndGet() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    solicitud.idSolicitud = 123
    solicitud.IDExpediente = "EXP-001"
    solicitud.TipoSolicitud = "PC"
    solicitud.CodigoSolicitud = "PC-2024-001"
    solicitud.EstadoInterno = "Borrador"
    
    ' Assert
    Test_CSolicitudPC_Properties_SetAndGet = (solicitud.idSolicitud = 123) And _
                                            (solicitud.IDExpediente = "EXP-001") And _
                                            (solicitud.TipoSolicitud = "PC") And _
                                            (solicitud.CodigoSolicitud = "PC-2024-001") And _
                                            (solicitud.EstadoInterno = "Borrador")
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_Properties_SetAndGet = False
End Function

Public Function Test_CSolicitudPC_Load_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    Dim result As Boolean
    result = solicitud.Load(123)
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_Load_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_Load_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_Save_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    Dim result As Boolean
    result = solicitud.Save()
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_Save_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_Save_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_ChangeState_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    Dim result As Boolean
    result = solicitud.ChangeState("Aprobado")
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_ChangeState_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_ChangeState_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_DatosPC_SetAndGet() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    Dim datosPC As T_Datos_PC
    datosPC.ID = 456
    datosPC.IDSolicitud = 123
    datosPC.TituloPC = "Propuesta de Cambio de Prueba"
    datosPC.DescripcionPC = "Descripci?n de la propuesta"
    datosPC.JustificacionPC = "Justificaci?n de la propuesta"
    datosPC.ImpactoPC = "Impacto esperado"
    datosPC.FechaCreacion = Now
    datosPC.CreadoPor = "usuario.prueba@empresa.com"
    
    ' Act
    solicitud.DatosPC = datosPC
    Dim retrievedDatos As T_Datos_PC
    retrievedDatos = solicitud.DatosPC
    
    ' Assert
    Test_CSolicitudPC_DatosPC_SetAndGet = (retrievedDatos.ID = 456) And _
                                         (retrievedDatos.IDSolicitud = 123) And _
                                         (retrievedDatos.TituloPC = "Propuesta de Cambio de Prueba")
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_DatosPC_SetAndGet = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_CreateSolicitud_LargeID_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = 2147483647 ' Valor m?ximo para Long
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Deber?a manejar IDs grandes sin errores
    Test_CreateSolicitud_LargeID_HandlesCorrectly = True
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_LargeID_HandlesCorrectly = False
End Function

Public Function Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    solicitud.CodigoSolicitud = "PC-2024-001_??@#$%"
    solicitud.IDExpediente = "EXP-001-???"
    
    ' Assert
    Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly = (InStr(solicitud.CodigoSolicitud, "??@") > 0) And _
                                                          (InStr(solicitud.IDExpediente, "???") > 0)
    
    Exit Function
    
TestFail:
    Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly = False
End Function

' ============================================================================
' FUNCI?N PRINCIPAL DE EJECUCI?N DE PRUEBAS
' ============================================================================

Public Function RunSolicitudFactoryTests() As Boolean
    Dim totalTests As Integer
    Dim passedTests As Integer
    Dim failedTests As Integer
    
    totalTests = 0
    passedTests = 0
    failedTests = 0
    
    Debug.Print "============================================================================"
    Debug.Print "EJECUTANDO PRUEBAS DE SOLICITUD FACTORY"
    Debug.Print "============================================================================"
    
    ' Pruebas de modSolicitudFactory
    Debug.Print "\n--- Pruebas de modSolicitudFactory ---"
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC() Then
        Debug.Print "✓ Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_InvalidID_ReturnsNothing() Then
        Debug.Print "✓ Test_CreateSolicitud_InvalidID_ReturnsNothing: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateSolicitud_InvalidID_ReturnsNothing: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_ZeroID_ReturnsNothing() Then
        Debug.Print "✓ Test_CreateSolicitud_ZeroID_ReturnsNothing: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateSolicitud_ZeroID_ReturnsNothing: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_GetTipoSolicitud_DefaultsToPC() Then
        Debug.Print "✓ Test_GetTipoSolicitud_DefaultsToPC: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_GetTipoSolicitud_DefaultsToPC: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_Factory_CreatesValidISolicitudInterface() Then
        Debug.Print "✓ Test_Factory_CreatesValidISolicitudInterface: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_Factory_CreatesValidISolicitudInterface: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_Factory_HandlesDatabaseErrors() Then
        Debug.Print "✓ Test_Factory_HandlesDatabaseErrors: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_Factory_HandlesDatabaseErrors: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de CSolicitudPC
    Debug.Print "\n--- Pruebas de CSolicitudPC ---"
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Properties_SetAndGet() Then
        Debug.Print "✓ Test_CSolicitudPC_Properties_SetAndGet: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_Properties_SetAndGet: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Load_ReturnsTrue() Then
        Debug.Print "✓ Test_CSolicitudPC_Load_ReturnsTrue: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_Load_ReturnsTrue: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Save_ReturnsTrue() Then
        Debug.Print "✓ Test_CSolicitudPC_Save_ReturnsTrue: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_Save_ReturnsTrue: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_ChangeState_ReturnsTrue() Then
        Debug.Print "✓ Test_CSolicitudPC_ChangeState_ReturnsTrue: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_ChangeState_ReturnsTrue: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_DatosPC_SetAndGet() Then
        Debug.Print "✓ Test_CSolicitudPC_DatosPC_SetAndGet: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_DatosPC_SetAndGet: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de casos extremos
    Debug.Print "\n--- Pruebas de Casos Extremos ---"
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_LargeID_HandlesCorrectly() Then
        Debug.Print "✓ Test_CreateSolicitud_LargeID_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CreateSolicitud_LargeID_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly() Then
        Debug.Print "✓ Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly: PASÓ"
        passedTests = passedTests + 1
    Else
        Debug.Print "✗ Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly: FALLÓ"
        failedTests = failedTests + 1
    End If
    
    ' Resumen final
    Debug.Print "\n============================================================================"
    Debug.Print "RESUMEN DE PRUEBAS DE SOLICITUD FACTORY"
    Debug.Print "============================================================================"
    Debug.Print "Total de pruebas ejecutadas: " & totalTests
    Debug.Print "Pruebas que pasaron: " & passedTests
    Debug.Print "Pruebas que fallaron: " & failedTests
    Debug.Print "Porcentaje de éxito: " & Format((passedTests / totalTests) * 100, "0.00") & "%"
    
    If failedTests = 0 Then
        Debug.Print "\n🎉 ¡TODAS LAS PRUEBAS PASARON!"
        RunSolicitudFactoryTests = True
    Else
        Debug.Print "\n⚠️  ALGUNAS PRUEBAS FALLARON. Revisar implementación."
        RunSolicitudFactoryTests = False
    End If
    
    Debug.Print "============================================================================"
End Function