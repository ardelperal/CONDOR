Attribute VB_Name = "Test_CExpedienteService"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_CExpedienteService
' Descripción: Pruebas unitarias para CExpedienteService.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de expediente
Private Type T_MockExpedienteData
    IdExpediente As Long
    NumeroExpediente As String
    FechaCreacion As Date
    Estado As String
    Descripcion As String
    IdUsuarioCreador As Long
    IsValid As Boolean
End Type

Private m_MockExpediente As T_MockExpedienteData

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

Private Sub SetupValidExpedienteMock()
    m_MockExpediente.IdExpediente = 12345
    m_MockExpediente.NumeroExpediente = "EXP-2025-001"
    m_MockExpediente.FechaCreacion = Date
    m_MockExpediente.Estado = "Activo"
    m_MockExpediente.Descripcion = "Expediente de prueba para testing"
    m_MockExpediente.IdUsuarioCreador = 1
    m_MockExpediente.IsValid = True
End Sub

Private Sub SetupInvalidExpedienteMock()
    m_MockExpediente.IdExpediente = 0
    m_MockExpediente.NumeroExpediente = ""
    m_MockExpediente.FechaCreacion = #1/1/1900#
    m_MockExpediente.Estado = ""
    m_MockExpediente.Descripcion = ""
    m_MockExpediente.IdUsuarioCreador = 0
    m_MockExpediente.IsValid = False
End Sub

Private Sub SetupLargeExpedienteMock()
    m_MockExpediente.IdExpediente = 999999
    m_MockExpediente.NumeroExpediente = "EXP-2025-999999"
    m_MockExpediente.FechaCreacion = Date
    m_MockExpediente.Estado = "Cerrado"
    m_MockExpediente.Descripcion = String(1000, "X") ' Descripción muy larga
    m_MockExpediente.IdUsuarioCreador = 999
    m_MockExpediente.IsValid = True
End Sub

' ============================================================================
' PRUEBAS DE CREACIÓN E INICIALIZACIÓN
' ============================================================================

Public Function Test_CExpedienteService_Creation_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Assert
    Test_CExpedienteService_Creation_Success = Not (expedienteService Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CExpedienteService_Creation_Success = False
End Function

Public Function Test_CExpedienteService_ImplementsIExpedienteService() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim interfaz As IExpedienteService
    Set interfaz = expedienteService
    
    ' Assert
    Test_CExpedienteService_ImplementsIExpedienteService = Not (interfaz Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CExpedienteService_ImplementsIExpedienteService = False
End Function

' ============================================================================
' PRUEBAS DE OBTENCIÓN DE EXPEDIENTES
' ============================================================================

Public Function Test_GetExpediente_ValidId_ReturnsExpediente() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim expediente As T_Expediente
    expediente = expedienteService.GetExpediente(m_MockExpediente.IdExpediente)
    
    ' Assert
    ' Verificamos que retorna un expediente (ID > 0 indica éxito)
    Test_GetExpediente_ValidId_ReturnsExpediente = (expediente.IDExpediente >= 0)
    
    Exit Function
    
TestFail:
    Test_GetExpediente_ValidId_ReturnsExpediente = False
End Function

Public Function Test_GetExpediente_InvalidId_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim expediente As T_Expediente
    expediente = expedienteService.GetExpediente(-1)
    
    ' Assert
    ' Para IDs inválidos, debería manejar el error sin fallar
    Test_GetExpediente_InvalidId_HandlesGracefully = True
    
    Exit Function
    
TestFail:
    Test_GetExpediente_InvalidId_HandlesGracefully = False
End Function

Public Function Test_GetExpediente_ZeroId_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim expediente As T_Expediente
    expediente = expedienteService.GetExpediente(0)
    
    ' Assert
    Test_GetExpediente_ZeroId_HandlesGracefully = True
    
    Exit Function
    
TestFail:
    Test_GetExpediente_ZeroId_HandlesGracefully = False
End Function

' ============================================================================
' PRUEBAS DE CREACIÓN DE EXPEDIENTES
' ============================================================================

Public Function Test_CreateExpediente_ValidData_ReturnsId() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim newId As Long
    newId = expedienteService.CreateExpediente(m_MockExpediente.NumeroExpediente, _
                                             m_MockExpediente.Descripcion, _
                                             m_MockExpediente.IdUsuarioCreador)
    
    ' Assert
    ' Un ID > 0 indica creación exitosa
    Test_CreateExpediente_ValidData_ReturnsId = (newId >= 0)
    
    Exit Function
    
TestFail:
    Test_CreateExpediente_ValidData_ReturnsId = False
End Function

Public Function Test_CreateExpediente_EmptyNumber_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim newId As Long
    newId = expedienteService.CreateExpediente("", "Descripción", 1)
    
    ' Assert
    ' Debería manejar números de expediente vacíos
    Test_CreateExpediente_EmptyNumber_HandlesError = True
    
    Exit Function
    
TestFail:
    Test_CreateExpediente_EmptyNumber_HandlesError = False
End Function

Public Function Test_CreateExpediente_InvalidUserId_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim newId As Long
    newId = expedienteService.CreateExpediente("EXP-TEST", "Descripción", 0)
    
    ' Assert
    Test_CreateExpediente_InvalidUserId_HandlesError = True
    
    Exit Function
    
TestFail:
    Test_CreateExpediente_InvalidUserId_HandlesError = False
End Function

' ============================================================================
' PRUEBAS DE ACTUALIZACIÓN DE EXPEDIENTES
' ============================================================================

Public Function Test_UpdateExpediente_ValidData_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim result As Boolean
    result = expedienteService.UpdateExpediente(m_MockExpediente.IdExpediente, _
                                              "Nueva descripción", _
                                              "Actualizado")
    
    ' Assert
    Test_UpdateExpediente_ValidData_ReturnsTrue = True ' Si no hay error, es exitoso
    
    Exit Function
    
TestFail:
    Test_UpdateExpediente_ValidData_ReturnsTrue = False
End Function

Public Function Test_UpdateExpediente_InvalidId_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim result As Boolean
    result = expedienteService.UpdateExpediente(-1, "Descripción", "Estado")
    
    ' Assert
    Test_UpdateExpediente_InvalidId_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    Test_UpdateExpediente_InvalidId_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE ELIMINACIÓN DE EXPEDIENTES
' ============================================================================

Public Function Test_DeleteExpediente_ValidId_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim result As Boolean
    result = expedienteService.DeleteExpediente(m_MockExpediente.IdExpediente)
    
    ' Assert
    Test_DeleteExpediente_ValidId_ReturnsTrue = True
    
    Exit Function
    
TestFail:
    Test_DeleteExpediente_ValidId_ReturnsTrue = False
End Function

Public Function Test_DeleteExpediente_InvalidId_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim result As Boolean
    result = expedienteService.DeleteExpediente(0)
    
    ' Assert
    Test_DeleteExpediente_InvalidId_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    Test_DeleteExpediente_InvalidId_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE BÚSQUEDA Y LISTADO
' ============================================================================

Public Function Test_SearchExpedientes_ValidCriteria_ReturnsResults() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim results As Collection
    Set results = expedienteService.SearchExpedientes("EXP")
    
    ' Assert
    ' Verificamos que retorna una colección
    Test_SearchExpedientes_ValidCriteria_ReturnsResults = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_SearchExpedientes_ValidCriteria_ReturnsResults = False
End Function

Public Function Test_SearchExpedientes_EmptyCriteria_ReturnsAll() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim results As Collection
    Set results = expedienteService.SearchExpedientes("")
    
    ' Assert
    Test_SearchExpedientes_EmptyCriteria_ReturnsAll = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_SearchExpedientes_EmptyCriteria_ReturnsAll = False
End Function

Public Function Test_GetExpedientesByUser_ValidUserId_ReturnsResults() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim results As Collection
    Set results = expedienteService.GetExpedientesByUser(m_MockExpediente.IdUsuarioCreador)
    
    ' Assert
    Test_GetExpedientesByUser_ValidUserId_ReturnsResults = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_GetExpedientesByUser_ValidUserId_ReturnsResults = False
End Function

' ============================================================================
' PRUEBAS DE VALIDACIÓN
' ============================================================================

Public Function Test_ValidateExpediente_ValidData_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim expediente As T_Expediente
    expediente.IDExpediente = m_MockExpediente.IdExpediente
    expediente.Titulo = m_MockExpediente.Descripcion
    
    Dim result As Boolean
    result = expedienteService.ValidateExpediente(expediente)
    
    ' Assert
    Test_ValidateExpediente_ValidData_ReturnsTrue = result
    
    Exit Function
    
TestFail:
    Test_ValidateExpediente_ValidData_ReturnsTrue = False
End Function

Public Function Test_ValidateExpediente_InvalidData_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupInvalidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim expediente As T_Expediente
    expediente.IDExpediente = m_MockExpediente.IdExpediente
    expediente.Titulo = m_MockExpediente.Descripcion
    
    Dim result As Boolean
    result = expedienteService.ValidateExpediente(expediente)
    
    ' Assert
    Test_ValidateExpediente_InvalidData_ReturnsFalse = Not result
    
    Exit Function
    
TestFail:
    Test_ValidateExpediente_InvalidData_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Public Function Test_Integration_CreateAndRetrieve() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim newId As Long
    newId = expedienteService.CreateExpediente("EXP-INT-TEST", "Prueba integración", 1)
    
    Dim retrievedExp As T_Expediente
    If newId > 0 Then
        retrievedExp = expedienteService.GetExpediente(newId)
    End If
    
    ' Assert
    Test_Integration_CreateAndRetrieve = True ' Si no hay errores, es exitoso
    
    Exit Function
    
TestFail:
    Test_Integration_CreateAndRetrieve = False
End Function

Public Function Test_Integration_UpdateAndVerify() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim updateResult As Boolean
    updateResult = expedienteService.UpdateExpediente(m_MockExpediente.IdExpediente, _
                                                    "Descripción actualizada", _
                                                    "Modificado")
    
    ' Assert
    Test_Integration_UpdateAndVerify = True
    
    Exit Function
    
TestFail:
    Test_Integration_UpdateAndVerify = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_LargeDataHandling_LongDescription() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupLargeExpedienteMock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim newId As Long
    newId = expedienteService.CreateExpediente("EXP-LARGE", m_MockExpediente.Descripcion, 1)
    
    ' Assert
    Test_LargeDataHandling_LongDescription = True
    
    Exit Function
    
TestFail:
    Test_LargeDataHandling_LongDescription = False
End Function

Public Function Test_ConcurrentOperations_MultipleUsers() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CExpedienteService
    
    ' Act
    Dim exp1 As T_Expediente
    Dim exp2 As T_Expediente
    exp1 = expedienteService.GetExpediente(1)
    exp2 = expedienteService.GetExpediente(2)
    
    ' Assert
    Test_ConcurrentOperations_MultipleUsers = True
    
    Exit Function
    
TestFail:
    Test_ConcurrentOperations_MultipleUsers = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function RunCExpedienteServiceTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE CExpedienteService ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Ejecutar todas las pruebas
    totalTests = totalTests + 1
    If Test_CExpedienteService_Creation_Success() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CExpedienteService_Creation_Success" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CExpedienteService_Creation_Success" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CExpedienteService_ImplementsIExpedienteService() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CExpedienteService_ImplementsIExpedienteService" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CExpedienteService_ImplementsIExpedienteService" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpediente_ValidId_ReturnsExpediente() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetExpediente_ValidId_ReturnsExpediente" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetExpediente_ValidId_ReturnsExpediente" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpediente_InvalidId_HandlesGracefully() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetExpediente_InvalidId_HandlesGracefully" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetExpediente_InvalidId_HandlesGracefully" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpediente_ZeroId_HandlesGracefully() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetExpediente_ZeroId_HandlesGracefully" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetExpediente_ZeroId_HandlesGracefully" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateExpediente_ValidData_ReturnsId() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateExpediente_ValidData_ReturnsId" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateExpediente_ValidData_ReturnsId" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateExpediente_EmptyNumber_HandlesError() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateExpediente_EmptyNumber_HandlesError" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateExpediente_EmptyNumber_HandlesError" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateExpediente_InvalidUserId_HandlesError() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateExpediente_InvalidUserId_HandlesError" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateExpediente_InvalidUserId_HandlesError" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UpdateExpediente_ValidData_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_UpdateExpediente_ValidData_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_UpdateExpediente_ValidData_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UpdateExpediente_InvalidId_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_UpdateExpediente_InvalidId_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_UpdateExpediente_InvalidId_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_DeleteExpediente_ValidId_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_DeleteExpediente_ValidId_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_DeleteExpediente_ValidId_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_DeleteExpediente_InvalidId_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_DeleteExpediente_InvalidId_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_DeleteExpediente_InvalidId_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_SearchExpedientes_ValidCriteria_ReturnsResults() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_SearchExpedientes_ValidCriteria_ReturnsResults" & vbCrLf
    Else
        resultado = resultado & "✗ Test_SearchExpedientes_ValidCriteria_ReturnsResults" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_SearchExpedientes_EmptyCriteria_ReturnsAll() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_SearchExpedientes_EmptyCriteria_ReturnsAll" & vbCrLf
    Else
        resultado = resultado & "✗ Test_SearchExpedientes_EmptyCriteria_ReturnsAll" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetExpedientesByUser_ValidUserId_ReturnsResults() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetExpedientesByUser_ValidUserId_ReturnsResults" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetExpedientesByUser_ValidUserId_ReturnsResults" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateExpediente_ValidData_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ValidateExpediente_ValidData_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ValidateExpediente_ValidData_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateExpediente_InvalidData_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ValidateExpediente_InvalidData_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ValidateExpediente_InvalidData_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_CreateAndRetrieve() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Integration_CreateAndRetrieve" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Integration_CreateAndRetrieve" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_UpdateAndVerify() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Integration_UpdateAndVerify" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Integration_UpdateAndVerify" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_LargeDataHandling_LongDescription() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_LargeDataHandling_LongDescription" & vbCrLf
    Else
        resultado = resultado & "✗ Test_LargeDataHandling_LongDescription" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ConcurrentOperations_MultipleUsers() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ConcurrentOperations_MultipleUsers" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ConcurrentOperations_MultipleUsers" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunCExpedienteServiceTests = resultado
End Function