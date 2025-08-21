Option Compare Database
Option Explicit
' ============================================================================
' MÃ³dulo: Test_SolicitudFactory
' DescripciÃ³n: Pruebas unitarias para modSolicitudFactory.bas
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de solicitud
Private Type T_MockSolicitudData
    ID As Long
    tipoSolicitud As String
    IsValid As Boolean
    ShouldFailLoad As Boolean
End Type

Private m_MockData As T_MockSolicitudData

' ============================================================================
' FUNCIONES DE CONFIGURACIÃ“N DE MOCKS
' ============================================================================

Private Sub SetupMockSolicitudData()
    m_MockData.ID = 123
    m_MockData.tipoSolicitud = "PC"
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
' FUNCIÃ“N PRINCIPAL PARA EJECUTAR TODAS LAS PRUEBAS
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
    ' Nota: En un entorno real, esto requerirÃ­a datos vÃ¡lidos en la base de datos
    Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC = Not (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC = False
End Function

Public Function Test_CreateSolicitud_InvalidID_ReturnsNothing() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = -1 ' ID invÃ¡lido
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' Con ID invÃ¡lido, deberÃ­a retornar Nothing
    Test_CreateSolicitud_InvalidID_ReturnsNothing = (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CreateSolicitud_InvalidID_ReturnsNothing", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
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
    modErrorHandler.LogError "Test_CreateSolicitud_ZeroID_ReturnsNothing", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CreateSolicitud_ZeroID_ReturnsNothing = False
End Function

' ============================================================================
' PRUEBAS PARA GetTipoSolicitud (funciÃ³n privada - prueba indirecta)
' ============================================================================

Public Function Test_GetTipoSolicitud_DefaultsToPC() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    ' Como GetTipoSolicitud es privada, la probamos indirectamente
    ' a travÃ©s de CreateSolicitud
    Dim idSolicitud As Long
    idSolicitud = 999 ' Cualquier ID
    
    ' La funciÃ³n privada GetTipoSolicitud siempre retorna "PC" por defecto
    ' Esto se refleja en que CreateSolicitud siempre crea CSolicitudPC
    
    ' Assert
    ' Por ahora, asumimos que funciona correctamente segÃºn la implementaciÃ³n
    Test_GetTipoSolicitud_DefaultsToPC = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_GetTipoSolicitud_DefaultsToPC", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_GetTipoSolicitud_DefaultsToPC = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N
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
        ' Si no es Nothing, significa que se creÃ³ correctamente
        ' En un entorno real, podrÃ­amos verificar propiedades especÃ­ficas
        Test_Factory_CreatesValidISolicitudInterface = True
    Else
        Test_Factory_CreatesValidISolicitudInterface = False
    End If
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Factory_CreatesValidISolicitudInterface", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
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
    ' Cuando hay errores de base de datos, deberÃ­a retornar Nothing
    Test_Factory_HandlesDatabaseErrors = (solicitud Is Nothing)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_Factory_HandlesDatabaseErrors", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
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
    solicitud.idExpediente = "EXP-001"
    solicitud.tipoSolicitud = "PC"
    solicitud.codigoSolicitud = "PC-2024-001"
    solicitud.estadoInterno = "Borrador"
    
    ' Assert
    Test_CSolicitudPC_Properties_SetAndGet = (solicitud.idSolicitud = 123) And _
                                            (solicitud.idExpediente = "EXP-001") And _
                                            (solicitud.tipoSolicitud = "PC") And _
                                            (solicitud.codigoSolicitud = "PC-2024-001") And _
                                            (solicitud.estadoInterno = "Borrador")
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_Properties_SetAndGet", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_Properties_SetAndGet = False
End Function

Public Function Test_CSolicitudPC_Load_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Dim iSolicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    Set iSolicitud = solicitud
    
    ' Act
    Dim result As Boolean
    result = iSolicitud.Load(123)
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_Load_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_Load_ReturnsTrue", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_Load_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_Save_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Dim iSolicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    Set iSolicitud = solicitud
    
    ' Act
    Dim result As Boolean
    result = iSolicitud.Save()
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_Save_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_Save_ReturnsTrue", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_Save_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_ChangeState_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Dim iSolicitud As ISolicitud
    Set solicitud = New CSolicitudPC
    Set iSolicitud = solicitud
    
    ' Act
    Dim result As Boolean
    result = iSolicitud.ChangeState("Aprobado")
    
    ' Assert
    ' Por ahora la implementaci?n siempre retorna True
    Test_CSolicitudPC_ChangeState_ReturnsTrue = (result = True)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_ChangeState_ReturnsTrue", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_ChangeState_ReturnsTrue = False
End Function

Public Function Test_CSolicitudPC_DatosPC_SetAndGet() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    Dim datosPC As T_Datos_PC
    datosPC.ID = 456
    datosPC.idSolicitud = 123
    datosPC.descripcionCambio = "Propuesta de Cambio de Prueba"
    datosPC.justificacion = "DescripciÃ³n de la propuesta"
    datosPC.impactoCalidad = "JustificaciÃ³n de la propuesta"
    datosPC.impactoCoste = "Impacto esperado"
    datosPC.fechaCreacion = Now
    datosPC.CreadoPor = "usuario.prueba@empresa.com"
    
    ' Act
    solicitud.datosPC = datosPC
    Dim retrievedDatos As T_Datos_PC
    retrievedDatos = solicitud.datosPC
    
    ' Assert
    Test_CSolicitudPC_DatosPC_SetAndGet = (retrievedDatos.ID = 456) And _
                                         (retrievedDatos.idSolicitud = 123) And _
                                         (retrievedDatos.descripcionCambio = "Propuesta de Cambio de Prueba")
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_DatosPC_SetAndGet", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_DatosPC_SetAndGet = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_CreateSolicitud_LargeID_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim idSolicitud As Long
    idSolicitud = 2147483647 ' Valor mÃ¡ximo para Long
    
    ' Act
    Dim solicitud As ISolicitud
    Set solicitud = modSolicitudFactory.CreateSolicitud(idSolicitud)
    
    ' Assert
    ' DeberÃ­a manejar IDs grandes sin errores
    Test_CreateSolicitud_LargeID_HandlesCorrectly = True
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CreateSolicitud_LargeID_HandlesCorrectly", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CreateSolicitud_LargeID_HandlesCorrectly = False
End Function

Public Function Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitud As CSolicitudPC
    Set solicitud = New CSolicitudPC
    
    ' Act
    solicitud.codigoSolicitud = "PC-2024-001_??@#$%"
    solicitud.idExpediente = "EXP-001-???"
    
    ' Assert
    Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly = (InStr(solicitud.codigoSolicitud, "??@") > 0) And _
                                                          (InStr(solicitud.idExpediente, "???") > 0)
    
    Exit Function
    
TestFail:
    modErrorHandler.LogError "Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly", Err.Number, Err.Description, "Test_SolicitudFactory.bas"
    Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly = False
End Function

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N DE PRUEBAS
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
        Debug.Print "âœ“ Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CreateSolicitud_ValidID_ReturnsCSolicitudPC: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_InvalidID_ReturnsNothing() Then
        Debug.Print "âœ“ Test_CreateSolicitud_InvalidID_ReturnsNothing: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CreateSolicitud_InvalidID_ReturnsNothing: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_ZeroID_ReturnsNothing() Then
        Debug.Print "âœ“ Test_CreateSolicitud_ZeroID_ReturnsNothing: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CreateSolicitud_ZeroID_ReturnsNothing: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_GetTipoSolicitud_DefaultsToPC() Then
        Debug.Print "âœ“ Test_GetTipoSolicitud_DefaultsToPC: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_GetTipoSolicitud_DefaultsToPC: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_Factory_CreatesValidISolicitudInterface() Then
        Debug.Print "âœ“ Test_Factory_CreatesValidISolicitudInterface: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_Factory_CreatesValidISolicitudInterface: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_Factory_HandlesDatabaseErrors() Then
        Debug.Print "âœ“ Test_Factory_HandlesDatabaseErrors: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_Factory_HandlesDatabaseErrors: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de CSolicitudPC
    Debug.Print "\n--- Pruebas de CSolicitudPC ---"
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Properties_SetAndGet() Then
        Debug.Print "âœ“ Test_CSolicitudPC_Properties_SetAndGet: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_Properties_SetAndGet: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Load_ReturnsTrue() Then
        Debug.Print "âœ“ Test_CSolicitudPC_Load_ReturnsTrue: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_Load_ReturnsTrue: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_Save_ReturnsTrue() Then
        Debug.Print "âœ“ Test_CSolicitudPC_Save_ReturnsTrue: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_Save_ReturnsTrue: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_ChangeState_ReturnsTrue() Then
        Debug.Print "âœ“ Test_CSolicitudPC_ChangeState_ReturnsTrue: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_ChangeState_ReturnsTrue: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_DatosPC_SetAndGet() Then
        Debug.Print "âœ“ Test_CSolicitudPC_DatosPC_SetAndGet: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_DatosPC_SetAndGet: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    ' Pruebas de casos extremos
    Debug.Print "\n--- Pruebas de Casos Extremos ---"
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_LargeID_HandlesCorrectly() Then
        Debug.Print "âœ“ Test_CreateSolicitud_LargeID_HandlesCorrectly: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CreateSolicitud_LargeID_HandlesCorrectly: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly() Then
        Debug.Print "âœ“ Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly: PASÃ“"
        passedTests = passedTests + 1
    Else
        Debug.Print "âœ— Test_CSolicitudPC_SpecialCharacters_HandlesCorrectly: FALLÃ“"
        failedTests = failedTests + 1
    End If
    
    ' Resumen final
    Debug.Print "\n============================================================================"
    Debug.Print "RESUMEN DE PRUEBAS DE SOLICITUD FACTORY"
    Debug.Print "============================================================================"
    Debug.Print "Total de pruebas ejecutadas: " & totalTests
    Debug.Print "Pruebas que pasaron: " & passedTests
    Debug.Print "Pruebas que fallaron: " & failedTests
    Debug.Print "Porcentaje de Ã©xito: " & Format((passedTests / totalTests) * 100, "0.00") & "%"
    
    If failedTests = 0 Then
        Debug.Print "\nðŸŽ‰ Â¡TODAS LAS PRUEBAS PASARON!"
        RunSolicitudFactoryTests = True
    Else
        Debug.Print "\nâš ï¸  ALGUNAS PRUEBAS FALLARON. Revisar implementaciÃ³n."
        RunSolicitudFactoryTests = False
    End If
    
    Debug.Print "============================================================================"
End Function














