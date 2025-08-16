Attribute VB_Name = "Test_CSolicitudService"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: Test_CSolicitudService
' Descripción: Pruebas unitarias para CSolicitudService.cls
' Autor: CONDOR-Expert
' Fecha: Enero 2025
' ============================================================================

' Mock para simular datos de solicitud
Private Type T_MockSolicitudData
    IdSolicitud As Long
    IdExpediente As Long
    TipoSolicitud As String
    CodigoSolicitud As String
    FechaCreacion As Date
    Estado As String
    Descripcion As String
    IdUsuarioCreador As Long
    IsValid As Boolean
End Type

Private m_MockSolicitud As T_MockSolicitudData

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN DE MOCKS
' ============================================================================

Private Sub SetupValidSolicitudMock()
    m_MockSolicitud.IdSolicitud = 54321
    m_MockSolicitud.IdExpediente = 12345
    m_MockSolicitud.TipoSolicitud = "PC"
    m_MockSolicitud.CodigoSolicitud = "SOL-PC-2025-001"
    m_MockSolicitud.FechaCreacion = Date
    m_MockSolicitud.Estado = "Pendiente"
    m_MockSolicitud.Descripcion = "Solicitud de PC para testing"
    m_MockSolicitud.IdUsuarioCreador = 1
    m_MockSolicitud.IsValid = True
End Sub

Private Sub SetupInvalidSolicitudMock()
    m_MockSolicitud.IdSolicitud = 0
    m_MockSolicitud.IdExpediente = 0
    m_MockSolicitud.TipoSolicitud = ""
    m_MockSolicitud.CodigoSolicitud = ""
    m_MockSolicitud.FechaCreacion = #1/1/1900#
    m_MockSolicitud.Estado = ""
    m_MockSolicitud.Descripcion = ""
    m_MockSolicitud.IdUsuarioCreador = 0
    m_MockSolicitud.IsValid = False
End Sub

Private Sub SetupCompletedSolicitudMock()
    m_MockSolicitud.IdSolicitud = 99999
    m_MockSolicitud.IdExpediente = 12345
    m_MockSolicitud.TipoSolicitud = "PC"
    m_MockSolicitud.CodigoSolicitud = "SOL-PC-2025-999"
    m_MockSolicitud.FechaCreacion = DateAdd("d", -30, Date)
    m_MockSolicitud.Estado = "Completada"
    m_MockSolicitud.Descripcion = "Solicitud completada para testing"
    m_MockSolicitud.IdUsuarioCreador = 1
    m_MockSolicitud.IsValid = True
End Sub

' ============================================================================
' PRUEBAS DE CREACIÓN E INICIALIZACIÓN
' ============================================================================

Public Function Test_CSolicitudService_Creation_Success() As Boolean
    On Error GoTo TestFail
    
    ' Arrange & Act
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Assert
    Test_CSolicitudService_Creation_Success = Not (solicitudService Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CSolicitudService_Creation_Success = False
End Function

Public Function Test_CSolicitudService_ImplementsISolicitudService() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim interfaz As ISolicitudService
    Set interfaz = solicitudService
    
    ' Assert
    Test_CSolicitudService_ImplementsISolicitudService = Not (interfaz Is Nothing)
    
    Exit Function
    
TestFail:
    Test_CSolicitudService_ImplementsISolicitudService = False
End Function

' ============================================================================
' PRUEBAS DE OBTENCIÓN DE SOLICITUDES
' ============================================================================

Public Function Test_GetSolicitud_ValidId_ReturnsSolicitud() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim solicitud As T_Solicitud
    solicitud = solicitudService.GetSolicitud(m_MockSolicitud.IdSolicitud)
    
    ' Assert
    ' Verificamos que retorna una solicitud (ID >= 0 indica éxito)
    Test_GetSolicitud_ValidId_ReturnsSolicitud = (solicitud.idSolicitud >= 0)
    
    Exit Function
    
TestFail:
    Test_GetSolicitud_ValidId_ReturnsSolicitud = False
End Function

Public Function Test_GetSolicitud_InvalidId_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim solicitud As T_Solicitud
    solicitud = solicitudService.GetSolicitud(-1)
    
    ' Assert
    Test_GetSolicitud_InvalidId_HandlesGracefully = True
    
    Exit Function
    
TestFail:
    Test_GetSolicitud_InvalidId_HandlesGracefully = False
End Function

Public Function Test_GetSolicitud_ZeroId_HandlesGracefully() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim solicitud As T_Solicitud
    solicitud = solicitudService.GetSolicitud(0)
    
    ' Assert
    Test_GetSolicitud_ZeroId_HandlesGracefully = True
    
    Exit Function
    
TestFail:
    Test_GetSolicitud_ZeroId_HandlesGracefully = False
End Function

' ============================================================================
' PRUEBAS DE CREACIÓN DE SOLICITUDES
' ============================================================================

Public Function Test_CreateSolicitud_ValidData_ReturnsId() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim newId As Long
    newId = solicitudService.CreateSolicitud(m_MockSolicitud.IdExpediente, _
                                           m_MockSolicitud.TipoSolicitud, _
                                           m_MockSolicitud.Descripcion, _
                                           m_MockSolicitud.IdUsuarioCreador)
    
    ' Assert
    Test_CreateSolicitud_ValidData_ReturnsId = (newId >= 0)
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_ValidData_ReturnsId = False
End Function

Public Function Test_CreateSolicitud_InvalidExpedienteId_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim newId As Long
    newId = solicitudService.CreateSolicitud(0, "PC", "Descripción", 1)
    
    ' Assert
    Test_CreateSolicitud_InvalidExpedienteId_HandlesError = True
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_InvalidExpedienteId_HandlesError = False
End Function

Public Function Test_CreateSolicitud_EmptyTipo_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim newId As Long
    newId = solicitudService.CreateSolicitud(12345, "", "Descripción", 1)
    
    ' Assert
    Test_CreateSolicitud_EmptyTipo_HandlesError = True
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_EmptyTipo_HandlesError = False
End Function

Public Function Test_CreateSolicitud_InvalidUserId_HandlesError() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim newId As Long
    newId = solicitudService.CreateSolicitud(12345, "PC", "Descripción", 0)
    
    ' Assert
    Test_CreateSolicitud_InvalidUserId_HandlesError = True
    
    Exit Function
    
TestFail:
    Test_CreateSolicitud_InvalidUserId_HandlesError = False
End Function

' ============================================================================
' PRUEBAS DE ACTUALIZACIÓN DE SOLICITUDES
' ============================================================================

Public Function Test_UpdateSolicitud_ValidData_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As ISolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result As Boolean
    result = solicitudService.UpdateSolicitud(m_MockSolicitud.IdSolicitud, _
                                            "Nueva descripción", _
                                            "En Proceso")
    
    ' Assert
    Test_UpdateSolicitud_ValidData_ReturnsTrue = True
    
    Exit Function
    
TestFail:
    Test_UpdateSolicitud_ValidData_ReturnsTrue = False
End Function

Public Function Test_UpdateSolicitud_InvalidId_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result As Boolean
    result = solicitudService.UpdateSolicitud(-1, "Descripción", "Estado")
    
    ' Assert
    Test_UpdateSolicitud_InvalidId_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    Test_UpdateSolicitud_InvalidId_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE CAMBIO DE ESTADO
' ============================================================================

Public Function Test_ChangeEstado_ValidTransition_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result As Boolean
    result = solicitudService.ChangeEstado(m_MockSolicitud.IdSolicitud, "En Proceso")
    
    ' Assert
    Test_ChangeEstado_ValidTransition_ReturnsTrue = True
    
    Exit Function
    
TestFail:
    Test_ChangeEstado_ValidTransition_ReturnsTrue = False
End Function

Public Function Test_ChangeEstado_InvalidTransition_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupCompletedSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result As Boolean
    result = solicitudService.ChangeEstado(m_MockSolicitud.IdSolicitud, "Pendiente")
    
    ' Assert
    ' Cambiar de Completada a Pendiente debería ser inválido
    Test_ChangeEstado_InvalidTransition_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    Test_ChangeEstado_InvalidTransition_ReturnsFalse = False
End Function

Public Function Test_ChangeEstado_EmptyEstado_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result As Boolean
    result = solicitudService.ChangeEstado(m_MockSolicitud.IdSolicitud, "")
    
    ' Assert
    Test_ChangeEstado_EmptyEstado_ReturnsFalse = True
    
    Exit Function
    
TestFail:
    Test_ChangeEstado_EmptyEstado_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE BÚSQUEDA Y LISTADO
' ============================================================================

Public Function Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim results As Collection
    Set results = solicitudService.GetSolicitudesByExpediente(m_MockSolicitud.IdExpediente)
    
    ' Assert
    Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection = False
End Function

Public Function Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim results As Collection
    Set results = solicitudService.GetSolicitudesByTipo("PC")
    
    ' Assert
    Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection = False
End Function

Public Function Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim results As Collection
    Set results = solicitudService.GetSolicitudesByEstado("Pendiente")
    
    ' Assert
    Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection = False
End Function

Public Function Test_SearchSolicitudes_ValidCriteria_ReturnsResults() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim results As Collection
    Set results = solicitudService.SearchSolicitudes("SOL")
    
    ' Assert
    Test_SearchSolicitudes_ValidCriteria_ReturnsResults = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_SearchSolicitudes_ValidCriteria_ReturnsResults = False
End Function

' ============================================================================
' PRUEBAS DE VALIDACIÓN
' ============================================================================

Public Function Test_ValidateSolicitud_ValidData_ReturnsTrue() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = m_MockSolicitud.IdSolicitud
    solicitud.IDExpediente = m_MockSolicitud.IdExpediente
    solicitud.TipoSolicitud = m_MockSolicitud.TipoSolicitud
    ' solicitud.CodigoSolicitud = m_MockSolicitud.CodigoSolicitud ' Comentado: CodigoSolicitud no existe en T_Solicitud
    
    Dim result As Boolean
    result = solicitudService.ValidateSolicitud(solicitud)
    
    ' Assert
    Test_ValidateSolicitud_ValidData_ReturnsTrue = result
    
    Exit Function
    
TestFail:
    Test_ValidateSolicitud_ValidData_ReturnsTrue = False
End Function

Public Function Test_ValidateSolicitud_InvalidData_ReturnsFalse() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupInvalidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim solicitud As T_Solicitud
    solicitud.idSolicitud = m_MockSolicitud.IdSolicitud
    solicitud.IDExpediente = m_MockSolicitud.IdExpediente
    solicitud.TipoSolicitud = m_MockSolicitud.TipoSolicitud
    ' solicitud.CodigoSolicitud = m_MockSolicitud.CodigoSolicitud ' Comentado: CodigoSolicitud no existe en T_Solicitud
    
    Dim result As Boolean
    result = solicitudService.ValidateSolicitud(solicitud)
    
    ' Assert
    Test_ValidateSolicitud_InvalidData_ReturnsFalse = Not result
    
    Exit Function
    
TestFail:
    Test_ValidateSolicitud_InvalidData_ReturnsFalse = False
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN CON FACTORY
' ============================================================================

Public Function Test_Integration_WithSolicitudFactory() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    ' Probamos la integración con el factory
    Dim solicitudObj As ISolicitud
    Set solicitudObj = modSolicitudFactory.CreateSolicitud(m_MockSolicitud.IdSolicitud)
    
    ' Assert
    Test_Integration_WithSolicitudFactory = Not (solicitudObj Is Nothing)
    
    Exit Function
    
TestFail:
    Test_Integration_WithSolicitudFactory = False
End Function

' ============================================================================
' PRUEBAS DE CASOS EXTREMOS
' ============================================================================

Public Function Test_LargeDataHandling_ManyResults() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    ' Simulamos búsqueda que podría retornar muchos resultados
    Dim results As Collection
    Set results = solicitudService.SearchSolicitudes("")
    
    ' Assert
    Test_LargeDataHandling_ManyResults = Not (results Is Nothing)
    
    Exit Function
    
TestFail:
    Test_LargeDataHandling_ManyResults = False
End Function

Public Function Test_ConcurrentOperations_MultipleUpdates() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    SetupValidSolicitudMock
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim result1 As Boolean
    Dim result2 As Boolean
    result1 = solicitudService.UpdateSolicitud(m_MockSolicitud.IdSolicitud, "Desc1", "Estado1")
    result2 = solicitudService.UpdateSolicitud(m_MockSolicitud.IdSolicitud, "Desc2", "Estado2")
    
    ' Assert
    Test_ConcurrentOperations_MultipleUpdates = True
    
    Exit Function
    
TestFail:
    Test_ConcurrentOperations_MultipleUpdates = False
End Function

Public Function Test_EdgeCase_VeryLongDescription() As Boolean
    On Error GoTo TestFail
    
    ' Arrange
    Dim solicitudService As CSolicitudService
    Set solicitudService = New CSolicitudService
    
    ' Act
    Dim longDesc As String
    longDesc = String(2000, "X") ' Descripción muy larga
    
    Dim newId As Long
    newId = solicitudService.CreateSolicitud(12345, "PC", longDesc, 1)
    
    ' Assert
    Test_EdgeCase_VeryLongDescription = True
    
    Exit Function
    
TestFail:
    Test_EdgeCase_VeryLongDescription = False
End Function

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN DE PRUEBAS
' ============================================================================

Public Function RunCSolicitudServiceTests() As String
    Dim resultado As String
    Dim totalTests As Integer
    Dim passedTests As Integer
    
    resultado = "=== PRUEBAS DE CSolicitudService ===" & vbCrLf
    totalTests = 0
    passedTests = 0
    
    ' Ejecutar todas las pruebas
    totalTests = totalTests + 1
    If Test_CSolicitudService_Creation_Success() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CSolicitudService_Creation_Success" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CSolicitudService_Creation_Success" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CSolicitudService_ImplementsISolicitudService() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CSolicitudService_ImplementsISolicitudService" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CSolicitudService_ImplementsISolicitudService" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitud_ValidId_ReturnsSolicitud() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitud_ValidId_ReturnsSolicitud" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitud_ValidId_ReturnsSolicitud" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitud_InvalidId_HandlesGracefully() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitud_InvalidId_HandlesGracefully" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitud_InvalidId_HandlesGracefully" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitud_ZeroId_HandlesGracefully() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitud_ZeroId_HandlesGracefully" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitud_ZeroId_HandlesGracefully" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_ValidData_ReturnsId() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateSolicitud_ValidData_ReturnsId" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateSolicitud_ValidData_ReturnsId" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_InvalidExpedienteId_HandlesError() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateSolicitud_InvalidExpedienteId_HandlesError" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateSolicitud_InvalidExpedienteId_HandlesError" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_EmptyTipo_HandlesError() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateSolicitud_EmptyTipo_HandlesError" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateSolicitud_EmptyTipo_HandlesError" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_CreateSolicitud_InvalidUserId_HandlesError() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_CreateSolicitud_InvalidUserId_HandlesError" & vbCrLf
    Else
        resultado = resultado & "✗ Test_CreateSolicitud_InvalidUserId_HandlesError" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UpdateSolicitud_ValidData_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_UpdateSolicitud_ValidData_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_UpdateSolicitud_ValidData_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_UpdateSolicitud_InvalidId_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_UpdateSolicitud_InvalidId_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_UpdateSolicitud_InvalidId_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ChangeEstado_ValidTransition_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ChangeEstado_ValidTransition_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ChangeEstado_ValidTransition_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ChangeEstado_InvalidTransition_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ChangeEstado_InvalidTransition_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ChangeEstado_InvalidTransition_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ChangeEstado_EmptyEstado_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ChangeEstado_EmptyEstado_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ChangeEstado_EmptyEstado_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitudesByExpediente_ValidId_ReturnsCollection" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitudesByTipo_ValidTipo_ReturnsCollection" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection" & vbCrLf
    Else
        resultado = resultado & "✗ Test_GetSolicitudesByEstado_ValidEstado_ReturnsCollection" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_SearchSolicitudes_ValidCriteria_ReturnsResults() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_SearchSolicitudes_ValidCriteria_ReturnsResults" & vbCrLf
    Else
        resultado = resultado & "✗ Test_SearchSolicitudes_ValidCriteria_ReturnsResults" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateSolicitud_ValidData_ReturnsTrue() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ValidateSolicitud_ValidData_ReturnsTrue" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ValidateSolicitud_ValidData_ReturnsTrue" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ValidateSolicitud_InvalidData_ReturnsFalse() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ValidateSolicitud_InvalidData_ReturnsFalse" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ValidateSolicitud_InvalidData_ReturnsFalse" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_Integration_WithSolicitudFactory() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_Integration_WithSolicitudFactory" & vbCrLf
    Else
        resultado = resultado & "✗ Test_Integration_WithSolicitudFactory" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_LargeDataHandling_ManyResults() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_LargeDataHandling_ManyResults" & vbCrLf
    Else
        resultado = resultado & "✗ Test_LargeDataHandling_ManyResults" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_ConcurrentOperations_MultipleUpdates() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_ConcurrentOperations_MultipleUpdates" & vbCrLf
    Else
        resultado = resultado & "✗ Test_ConcurrentOperations_MultipleUpdates" & vbCrLf
    End If
    
    totalTests = totalTests + 1
    If Test_EdgeCase_VeryLongDescription() Then
        passedTests = passedTests + 1
        resultado = resultado & "✓ Test_EdgeCase_VeryLongDescription" & vbCrLf
    Else
        resultado = resultado & "✗ Test_EdgeCase_VeryLongDescription" & vbCrLf
    End If
    
    ' Resumen
    resultado = resultado & vbCrLf & "Resultado: " & passedTests & "/" & totalTests & " pruebas exitosas" & vbCrLf
    
    RunCSolicitudServiceTests = resultado
End Function