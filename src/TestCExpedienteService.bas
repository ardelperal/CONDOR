Attribute VB_Name = "TestCExpedienteService"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CExpedienteService
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CExpedienteService
' utilizando mocks para todas las dependencias externas.
' ============================================================================

' Función principal que ejecuta todas las pruebas del módulo
Public Function TestCExpedienteServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestCExpedienteService")
    
    ' Ejecutar todas las pruebas unitarias
    Call suiteResult.AddTestResult(TestGetExpedienteByIdSuccess())
    Call suiteResult.AddTestResult(TestGetExpedienteByIdNotFound())
    Call suiteResult.AddTestResult(TestGetExpedientesParaSelectorSuccess())
    Call suiteResult.AddTestResult(TestGetExpedientesParaSelectorEmptyResult())
    
    Set TestCExpedienteServiceRunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetExpedienteById
' ============================================================================

' Prueba que GetExpedienteById devuelve correctamente un expediente cuando existe
Private Function TestGetExpedienteByIdSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedienteById debe devolver un expediente cuando existe")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim mockLogger As New CMockOperationLogger
    mockLogger.Reset
    Dim mockRepository As New CMockExpedienteRepository
    mockRepository.Reset
    Dim mockErrorHandler As New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Crear objeto EExpediente mock con datos de prueba
    Dim mockExpediente As New EExpediente
    mockExpediente.idExpediente = 123
    mockExpediente.NumeroExpediente = "EXP-2024-001"
    mockExpediente.Titulo = "Expediente de Prueba"
    mockExpediente.Descripcion = "Descripción de prueba"
    mockExpediente.FechaCreacion = #1/15/2024#
    mockExpediente.Estado = "Activo"
    mockExpediente.IdUsuarioCreador = 1
    mockExpediente.NombreUsuarioCreador = "Juan Pérez"
    Call mockRepository.ConfigureObtenerExpedientePorId(mockExpediente)
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CMockExpedienteService
    expedienteService.Reset
    expedienteService.Reset
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As EExpediente
    Set Resultado = expedienteService.GetExpedienteById(123)
    
    ' Assert - Verificar resultados
    Call modAssert.AssertEquals(123, Resultado.idExpediente, "ID del expediente debe coincidir")
    Call modAssert.AssertEquals("EXP-2024-001", Resultado.NumeroExpediente, "Número del expediente debe coincidir")
    Call modAssert.AssertEquals("Expediente de Prueba", Resultado.Titulo, "Título del expediente debe coincidir")
    Call modAssert.AssertEquals("Descripción de prueba", Resultado.Descripcion, "Descripción del expediente debe coincidir")
    Call modAssert.AssertEquals("Activo", Resultado.Estado, "Estado del expediente debe coincidir")
    Call modAssert.AssertEquals(1, Resultado.IdUsuarioCreador, "ID del usuario creador debe coincidir")
    Call modAssert.AssertEquals("Juan Pérez", Resultado.NombreUsuarioCreador, "Nombre del usuario creador debe coincidir")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    ' Limpiar recursos
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockRepository = Nothing
    Set mockErrorHandler = Nothing
    Set expedienteService = Nothing
    Set Resultado = Nothing
    Set TestGetExpedienteByIdSuccess = testResult
End Function

' Prueba que GetExpedienteById maneja correctamente cuando no se encuentra el expediente
Private Function TestGetExpedienteByIdNotFound() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedienteById debe devolver un objeto vacío si no se encuentra")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks con recordset vacío
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim mockLogger As New CMockOperationLogger
    mockLogger.Reset
    Dim mockRepository As New CMockExpedienteRepository
    mockRepository.Reset
    Dim mockErrorHandler As New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Configurar mock para devolver Nothing (expediente no encontrado)
    Call mockRepository.ConfigureObtenerExpedientePorId(Nothing)
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CMockExpedienteService
    expedienteService.Reset
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As EExpediente
    Set Resultado = expedienteService.GetExpedienteById(999)
    
    ' Assert - Verificar que devuelve estructura vacía
    Call modAssert.AssertEquals(0, Resultado.idExpediente, "ID debe ser 0 para expediente no encontrado")
    Call modAssert.AssertEquals("", Resultado.NumeroExpediente, "Número debe estar vacío para expediente no encontrado")
    Call modAssert.AssertEquals("", Resultado.Titulo, "Título debe estar vacío para expediente no encontrado")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    ' Limpiar recursos
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockRepository = Nothing
    Set mockErrorHandler = Nothing
    Set expedienteService = Nothing
    Set Resultado = Nothing
    Set TestGetExpedienteByIdNotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetExpedientesParaSelector
' ============================================================================

' Prueba que GetExpedientesParaSelector devuelve correctamente una lista de expedientes
Private Function TestGetExpedientesParaSelectorSuccess() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedientesParaSelector debe devolver una lista de expedientes")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim mockLogger As New CMockOperationLogger
    mockLogger.Reset
    Dim mockRepository As New CMockExpedienteRepository
    mockRepository.Reset
    Dim mockErrorHandler As New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Crear diccionario mock con lista de expedientes
    Dim mockExpedientes As New Scripting.Dictionary
    mockExpedientes.CompareMode = TextCompare
    Dim exp1 As New EExpediente
    exp1.idExpediente = 1
    exp1.NumeroExpediente = "EXP-2024-001"
    exp1.Titulo = "Expediente Uno"
    exp1.Estado = "Activo"
    mockExpedientes.Add CStr(exp1.idExpediente), exp1
    
    Dim exp2 As New EExpediente
    exp2.idExpediente = 2
    exp2.NumeroExpediente = "EXP-2024-002"
    exp2.Titulo = "Expediente Dos"
    exp2.Estado = "En Proceso"
    mockExpedientes.Add CStr(exp2.idExpediente), exp2
    
    Call mockRepository.ConfigureObtenerExpedientesActivosParaSelector(mockExpedientes)
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CMockExpedienteService
    expedienteService.Reset
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As Object
    Set Resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve una colección válida
    Call modAssert.AssertNotNull(Resultado, "El resultado no debe ser Nothing")
    Call modAssert.AssertEquals("Dictionary", TypeName(Resultado), "El resultado debe ser un Dictionary")
    
    ' Verificar que el diccionario tiene elementos
    Dim dict As Scripting.Dictionary
    Set dict = Resultado
    Call modAssert.AssertEquals(2, dict.Count, "El diccionario debe tener 2 expedientes")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockRepository = Nothing
    Set mockErrorHandler = Nothing
    Set expedienteService = Nothing
    Set Resultado = Nothing
    Set col = Nothing
    Set mockExpedientes = Nothing
    Set exp1 = Nothing
    Set exp2 = Nothing
    Set TestGetExpedientesParaSelectorSuccess = testResult
End Function

' Prueba que GetExpedientesParaSelector maneja correctamente cuando no hay expedientes
Private Function TestGetExpedientesParaSelectorEmptyResult() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedientesParaSelector debe devolver una colección vacía si no hay datos")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks con colección vacía
    Dim mockConfig As New CMockConfig
    mockConfig.Reset
    Dim mockLogger As New CMockOperationLogger
    mockLogger.Reset
    Dim mockRepository As New CMockExpedienteRepository
    mockRepository.Reset
    Dim mockErrorHandler As New CMockErrorHandlerService
    mockErrorHandler.Reset
    
    ' Crear diccionario mock vacío
    Dim mockExpedientes As New Scripting.Dictionary
    mockExpedientes.CompareMode = TextCompare
    Call mockRepository.ConfigureObtenerExpedientesActivosParaSelector(mockExpedientes)
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As IExpedienteService
    Set expedienteService = New CMockExpedienteService
    expedienteService.Reset
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As Object
    Set Resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve una colección válida pero vacía
    Call modAssert.AssertNotNull(Resultado, "El resultado no debe ser Nothing")
    Call modAssert.AssertEquals("Dictionary", TypeName(Resultado), "El resultado debe ser un Dictionary")
    
    Dim dict As Scripting.Dictionary
    Set dict = Resultado
    Call modAssert.AssertEquals(0, dict.Count, "El diccionario debe estar vacío")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set mockConfig = Nothing
    Set mockLogger = Nothing
    Set mockRepository = Nothing
    Set mockErrorHandler = Nothing
    Set expedienteService = Nothing
    Set Resultado = Nothing
    Set col = Nothing
    Set mockExpedientes = Nothing
    Set TestGetExpedientesParaSelectorEmptyResult = testResult
End Function

' ============================================================================
' FUNCIONES AUXILIARES ELIMINADAS
' ============================================================================

' Las funciones auxiliares para crear recordsets mock han sido eliminadas.
' Ahora se usan objetos de dominio (EExpediente) directamente en los tests.