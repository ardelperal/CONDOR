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
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Crear recordset mock con datos de expediente
    Dim mockRecordset As DAO.recordset
    Set mockRecordset = CreateMockExpedienteRecordset(123, "EXP-2024-001", "Expediente de Prueba", "Descripción de prueba", #1/15/2024#, "Activo", 1, "Juan Pérez")
    Call mockRepository.SetObtenerExpedientePorIdReturnValue(mockRecordset)
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
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
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
    Set TestGetExpedienteByIdSuccess = testResult
End Function

' Prueba que GetExpedienteById maneja correctamente cuando no se encuentra el expediente
Private Function TestGetExpedienteByIdNotFound() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedienteById debe devolver un objeto vacío si no se encuentra")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks con recordset vacío
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Crear recordset mock vacío (EOF = True)
    Dim mockRecordset As DAO.recordset
    Set mockRecordset = CreateEmptyRecordset()
    mockRepository.SetObtenerExpedientePorIdReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
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
    If Not mockRecordset Is Nothing Then
        mockRecordset.Close
        Set mockRecordset = Nothing
    End If
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
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Crear recordset mock con lista de expedientes
    Dim mockRecordset As DAO.recordset
    Set mockRecordset = CreateMockExpedientesListRecordset()
    mockRepository.SetObtenerExpedientePorIdReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As Object
    Set Resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido
    Call modAssert.AssertNotNull(Resultado, "El resultado no debe ser Nothing")
    Call modAssert.AssertEquals("DAO.Recordset", TypeName(Resultado), "El resultado debe ser un recordset")
    
    ' Verificar que el recordset tiene registros
    Dim rs As DAO.recordset
    Set rs = Resultado
    Call modAssert.AssertFalse(rs.EOF, "El recordset debe tener registros")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set TestGetExpedientesParaSelectorSuccess = testResult
End Function

' Prueba que GetExpedientesParaSelector maneja correctamente cuando no hay expedientes
Private Function TestGetExpedientesParaSelectorEmptyResult() As CTestResult
    Dim testResult As New CTestResult
    Call testResult.Initialize("GetExpedientesParaSelector debe devolver un recordset vacío si no hay datos")
    
    On Error GoTo TestFail
    
    ' Arrange - Configurar mocks con recordset vacío
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    Dim mockRepository As New CMockExpedienteRepository
    Dim mockErrorHandler As New CMockErrorHandlerService
    
    ' Crear recordset mock vacío
    Dim mockRecordset As DAO.recordset
    Set mockRecordset = CreateEmptyRecordset()
    mockRepository.SetObtenerExpedientePorIdReturnValue mockRecordset
    
    ' Crear servicio con dependencias mock
    Dim expedienteService As New CExpedienteService
    Call expedienteService.Initialize(mockConfig, mockLogger, mockRepository, mockErrorHandler)
    
    ' Act - Ejecutar el método bajo prueba
    Dim Resultado As Object
    Set Resultado = expedienteService.GetExpedientesParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido pero vacío
    Call modAssert.AssertNotNull(Resultado, "El resultado no debe ser Nothing")
    
    Dim rs As DAO.recordset
    Set rs = Resultado
    Call modAssert.AssertTrue(rs.EOF, "El recordset debe estar vacío")
    
    testResult.Pass
    GoTo Cleanup
    
TestFail:
    Call testResult.Fail("Error inesperado: " & Err.Description)
    
Cleanup:
    Set TestGetExpedientesParaSelectorEmptyResult = testResult
End Function

' ============================================================================
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK CON DAO
' ============================================================================

' Crea un recordset DAO con datos de un expediente específico usando BD temporal
Private Function CreateMockExpedienteRecordset(idExp As Long, numero As String, Titulo As String, Descripcion As String, fecha As Date, Estado As String, idUsuario As Long, nombreUsuario As String) As DAO.recordset
    Dim db As DAO.Database
    Dim rs As DAO.recordset
    Dim tempDbPath As String
    
    ' Crear base de datos temporal
    tempDbPath = Environ("TEMP") & "\TestExpediente_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = DBEngine.CreateDatabase(tempDbPath, dbLangGeneral)
    
    ' Crear tabla de expedientes
    db.Execute "CREATE TABLE Expedientes (" & _
               "idExpediente LONG PRIMARY KEY, " & _
               "NumeroExpediente TEXT(50), " & _
               "Titulo TEXT(255), " & _
               "Descripcion MEMO, " & _
               "FechaCreacion DATETIME, " & _
               "Estado TEXT(50), " & _
               "IdUsuarioCreador LONG, " & _
               "NombreUsuarioCreador TEXT(255))"
    
    ' Insertar datos de prueba
    db.Execute "INSERT INTO Expedientes (idExpediente, NumeroExpediente, Titulo, Descripcion, FechaCreacion, Estado, IdUsuarioCreador, NombreUsuarioCreador) " & _
               "VALUES (" & idExp & ", '" & numero & "', '" & Titulo & "', '" & Descripcion & "', #" & Format(fecha, "mm/dd/yyyy") & "#, '" & Estado & "', " & idUsuario & ", '" & nombreUsuario & "')"
    
    ' Abrir recordset
    Set rs = db.OpenRecordset("SELECT * FROM Expedientes", dbOpenDynaset)
    
    Set CreateMockExpedienteRecordset = rs
End Function

' Crea un recordset DAO con lista de expedientes para selector usando BD temporal
Private Function CreateMockExpedientesListRecordset() As DAO.recordset
    Dim db As DAO.Database
    Dim rs As DAO.recordset
    Dim tempDbPath As String
    
    ' Crear base de datos temporal
    tempDbPath = Environ("TEMP") & "\TestExpedientesList_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = DBEngine.CreateDatabase(tempDbPath, dbLangGeneral)
    
    ' Crear tabla de expedientes
    db.Execute "CREATE TABLE Expedientes (" & _
               "idExpediente LONG PRIMARY KEY, " & _
               "NumeroExpediente TEXT(50), " & _
               "Titulo TEXT(255), " & _
               "Estado TEXT(50))"
    
    ' Insertar datos de prueba
    db.Execute "INSERT INTO Expedientes (idExpediente, NumeroExpediente, Titulo, Estado) " & _
               "VALUES (1, 'EXP-2024-001', 'Expediente Uno', 'Activo')"
    db.Execute "INSERT INTO Expedientes (idExpediente, NumeroExpediente, Titulo, Estado) " & _
               "VALUES (2, 'EXP-2024-002', 'Expediente Dos', 'En Proceso')"
    
    ' Abrir recordset
    Set rs = db.OpenRecordset("SELECT * FROM Expedientes", dbOpenDynaset)
    
    Set CreateMockExpedientesListRecordset = rs
End Function

' Crea un recordset DAO vacío para simular casos donde no se encuentran datos
Private Function CreateEmptyRecordset() As DAO.recordset
    Dim db As DAO.Database
    Dim rs As DAO.recordset
    Dim tempDbPath As String
    
    ' Crear base de datos temporal
    tempDbPath = Environ("TEMP") & "\TestEmpty_" & Format(Now, "yyyymmddhhnnss") & ".accdb"
    Set db = DBEngine.CreateDatabase(tempDbPath, dbLangGeneral)
    
    ' Crear tabla vacía
    db.Execute "CREATE TABLE EmptyTable (id LONG PRIMARY KEY)"
    
    ' Abrir recordset vacío
    Set rs = db.OpenRecordset("SELECT * FROM EmptyTable", dbOpenDynaset)
    
    Set CreateEmptyRecordset = rs
End Function