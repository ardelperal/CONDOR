Attribute VB_Name = "Test_CExpedienteService"
Option Compare Database
Option Explicit

#If DEV_MODE Then

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CExpedienteService
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CExpedienteService
' utilizando mocks para todas las dependencias externas.
' Sigue la Lección 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable

' Función principal que ejecuta todas las pruebas del módulo
Public Function Test_CExpedienteService_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CExpedienteService"
    
    ' Ejecutar todas las pruebas unitarias
    suiteResult.AddTestResult Test_GetExpedienteById_Success()
    suiteResult.AddTestResult Test_GetExpedienteById_NotFound()
    suiteResult.AddTestResult Test_GetExpedientesParaSelector_Success()
    suiteResult.AddTestResult Test_GetExpedientesParaSelector_EmptyResult()
    
    Set Test_CExpedienteService_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetExpedienteById
' ============================================================================

' Prueba que GetExpedienteById devuelve correctamente un expediente cuando existe
Private Function Test_GetExpedienteById_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetExpedienteById_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    
    ' Crear recordset mock con datos de expediente
    Dim mockRecordset As DAO.Recordset
    Set mockRecordset = CreateMockExpedienteRecordset(123, "EXP-2024-001", "Expediente de Prueba", "Descripción de prueba", #1/15/2024#, "Activo", 1, "Juan Pérez")
    mockRepository.SetObtenerExpedientePorIdReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    expedienteService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act - Ejecutar el método bajo prueba
    Dim resultado As T_Expediente
    resultado = expedienteService.GetExpedienteById(123)
    
    ' Assert - Verificar resultados
    modAssert.AssertEquals 123, resultado.idExpediente, "ID del expediente debe coincidir"
    modAssert.AssertEquals "EXP-2024-001", resultado.NumeroExpediente, "Número del expediente debe coincidir"
    modAssert.AssertEquals "Expediente de Prueba", resultado.Titulo, "Título del expediente debe coincidir"
    modAssert.AssertEquals "Descripción de prueba", resultado.Descripcion, "Descripción del expediente debe coincidir"
    modAssert.AssertEquals "Activo", resultado.Estado, "Estado del expediente debe coincidir"
    modAssert.AssertEquals 1, resultado.IdUsuarioCreador, "ID del usuario creador debe coincidir"
    modAssert.AssertEquals "Juan Pérez", resultado.NombreUsuarioCreador, "Nombre del usuario creador debe coincidir"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetExpedienteById_Success = testResult
End Function

' Prueba que GetExpedienteById maneja correctamente cuando no se encuentra el expediente
Private Function Test_GetExpedienteById_NotFound() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetExpedienteById_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks con recordset vacío
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    
    ' Crear recordset mock vacío (EOF = True)
    Dim mockRecordset As DAO.Recordset
    Set mockRecordset = CreateEmptyRecordset()
    mockRepository.SetObtenerExpedientePorIdReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    expedienteService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act - Ejecutar el método bajo prueba
    Dim resultado As T_Expediente
    resultado = expedienteService.GetExpedienteById(999)
    
    ' Assert - Verificar que devuelve estructura vacía
    modAssert.AssertEquals 0, resultado.idExpediente, "ID debe ser 0 para expediente no encontrado"
    modAssert.AssertEquals "", resultado.NumeroExpediente, "Número debe estar vacío para expediente no encontrado"
    modAssert.AssertEquals "", resultado.Titulo, "Título debe estar vacío para expediente no encontrado"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set Test_GetExpedienteById_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetExpedientesParaSelector
' ============================================================================

' Prueba que GetExpedientesParaSelector devuelve correctamente una lista de expedientes
Private Function Test_GetExpedientesParaSelector_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetExpedientesParaSelector_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    
    ' Crear recordset mock con lista de expedientes
    Dim mockRecordset As DAO.Recordset
    Set mockRecordset = CreateMockExpedientesListRecordset()
    mockRepository.SetObtenerExpedientesActivosParaSelectorReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    expedienteService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act - Ejecutar el método bajo prueba
    Dim resultado As Object
    Set resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNothing resultado, "El resultado no debe ser Nothing"
    modAssert.AssertEquals "DAO.Recordset", TypeName(resultado), "El resultado debe ser un recordset"
    
    ' Verificar que el recordset tiene registros
    Dim rs As DAO.Recordset
    Set rs = resultado
    modAssert.AssertFalse rs.EOF, "El recordset debe tener registros"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set Test_GetExpedientesParaSelector_Success = testResult
End Function

' Prueba que GetExpedientesParaSelector maneja correctamente cuando no hay expedientes
Private Function Test_GetExpedientesParaSelector_EmptyResult() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_GetExpedientesParaSelector_EmptyResult"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks con recordset vacío
    Dim configService As IConfig
    Set configService = modConfig.CreateConfigService()
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    
    ' Crear recordset mock vacío
    Dim mockRecordset As DAO.Recordset
    Set mockRecordset = CreateEmptyRecordset()
    mockRepository.SetObtenerExpedientesActivosParaSelectorReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    expedienteService.Initialize mockConfig, mockLogger, mockRepository
    
    ' Act - Ejecutar el método bajo prueba
    Dim resultado As Object
    Set resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido pero vacío
    modAssert.AssertNotNothing resultado, "El resultado no debe ser Nothing"
    
    Dim rs As DAO.Recordset
    Set rs = resultado
    modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    Set Test_GetExpedientesParaSelector_EmptyResult = testResult
End Function

' ============================================================================
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK
' ============================================================================

' Crea un recordset mock con datos de un expediente específico
Private Function CreateMockExpedienteRecordset(idExp As Long, numero As String, titulo As String, descripcion As String, fecha As Date, estado As String, idUsuario As Long, nombreUsuario As String) As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos del recordset
    rs.Fields.Append "idExpediente", 3 ' adInteger
    rs.Fields.Append "NumeroExpediente", 202, 50 ' adVarWChar
    rs.Fields.Append "Titulo", 202, 255 ' adVarWChar
    rs.Fields.Append "Descripcion", 203 ' adLongVarWChar
    rs.Fields.Append "FechaCreacion", 7 ' adDate
    rs.Fields.Append "Estado", 202, 50 ' adVarWChar
    rs.Fields.Append "IdUsuarioCreador", 3 ' adInteger
    rs.Fields.Append "NombreUsuarioCreador", 202, 255 ' adVarWChar
    
    ' Abrir recordset en memoria
    rs.Open
    
    ' Añadir fila de datos de prueba
    rs.AddNew
    rs.Fields("idExpediente").Value = idExp
    rs.Fields("NumeroExpediente").Value = numero
    rs.Fields("Titulo").Value = titulo
    rs.Fields("Descripcion").Value = descripcion
    rs.Fields("FechaCreacion").Value = fecha
    rs.Fields("Estado").Value = estado
    rs.Fields("IdUsuarioCreador").Value = idUsuario
    rs.Fields("NombreUsuarioCreador").Value = nombreUsuario
    rs.Update
    
    ' Mover al primer registro
    rs.MoveFirst
    
    Set CreateMockExpedienteRecordset = rs
End Function

' Crea un recordset mock con lista de expedientes para selector
Private Function CreateMockExpedientesListRecordset() As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos del recordset
    rs.Fields.Append "idExpediente", 3 ' adInteger
    rs.Fields.Append "NumeroExpediente", 202, 50 ' adVarWChar
    rs.Fields.Append "Titulo", 202, 255 ' adVarWChar
    rs.Fields.Append "Estado", 202, 50 ' adVarWChar
    
    ' Abrir recordset en memoria
    rs.Open
    
    ' Añadir primer registro de prueba
    rs.AddNew
    rs.Fields("idExpediente").Value = 1
    rs.Fields("NumeroExpediente").Value = "EXP-2024-001"
    rs.Fields("Titulo").Value = "Expediente Uno"
    rs.Fields("Estado").Value = "Activo"
    rs.Update
    
    ' Añadir segundo registro de prueba
    rs.AddNew
    rs.Fields("idExpediente").Value = 2
    rs.Fields("NumeroExpediente").Value = "EXP-2024-002"
    rs.Fields("Titulo").Value = "Expediente Dos"
    rs.Fields("Estado").Value = "En Proceso"
    rs.Update
    
    ' Mover al primer registro
    rs.MoveFirst
    
    Set CreateMockExpedientesListRecordset = rs
End Function

' Crea un recordset vacío para simular casos donde no se encuentran datos
Private Function CreateEmptyRecordset() As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir un campo básico para el recordset vacío
    rs.Fields.Append "id", 3 ' adInteger
    
    ' Abrir recordset en memoria (sin añadir registros, queda vacío)
    rs.Open
    
    Set CreateEmptyRecordset = rs
End Function

#End If














