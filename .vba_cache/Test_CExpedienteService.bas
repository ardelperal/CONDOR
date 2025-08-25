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
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK CON DAO
' ============================================================================

' Crea un recordset DAO con datos de un expediente específico usando BD temporal
Private Function CreateMockExpedienteRecordset(idExp As Long, numero As String, titulo As String, descripcion As String, fecha As Date, estado As String, idUsuario As Long, nombreUsuario As String) As DAO.Recordset
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
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
               "VALUES (" & idExp & ", '" & numero & "', '" & titulo & "', '" & descripcion & "', #" & Format(fecha, "mm/dd/yyyy") & "#, '" & estado & "', " & idUsuario & ", '" & nombreUsuario & "')"
    
    ' Abrir recordset
    Set rs = db.OpenRecordset("SELECT * FROM Expedientes", dbOpenDynaset)
    
    Set CreateMockExpedienteRecordset = rs
End Function

' Crea un recordset DAO con lista de expedientes para selector usando BD temporal
Private Function CreateMockExpedientesListRecordset() As DAO.Recordset
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
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
Private Function CreateEmptyRecordset() As DAO.Recordset
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
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

#End If














