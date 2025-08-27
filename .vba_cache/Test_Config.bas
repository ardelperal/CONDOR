Attribute VB_Name = "Test_Config"
Option Compare Database
Option Explicit

' Suite de pruebas para el módulo de configuración
' Prueba la lógica del factory modConfig.bas

Public Function Test_Config_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_Config"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult Test_CreateConfigService_EntornoLocal_CargaConfiguracionCorrectamente()
    suiteResult.AddTestResult Test_CreateConfigService_SinEntornoConfigurado_LanzaError()
    suiteResult.AddTestResult Test_CreateConfigService_ConEntornoInvalido_LanzaError()
    
    Set Test_Config_RunAll = suiteResult
End Function

' Prueba de éxito: Entorno LOCAL carga configuración correctamente
Private Function Test_CreateConfigService_EntornoLocal_CargaConfiguracionCorrectamente() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_CreateConfigService_EntornoLocal_CargaConfiguracionCorrectamente"
    
    On Error GoTo ErrorHandler
    
    ' Arrange: Crear base de datos mock con TbLocalConfig y entorno LOCAL
    Dim db As DAO.Database
    Set db = CreateMockDatabaseWithEntorno("LOCAL")
    
    ' Act & Assert: Verificar que no se lance error
    Dim config As IConfig
    Set config = modConfig.CreateConfigService(db)
    
    ' Verificar que se devuelve una instancia válida
    modAssert.AssertNotNothing config, "El factory debe devolver una instancia de IConfig"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set Test_CreateConfigService_EntornoLocal_CargaConfiguracionCorrectamente = testResult
End Function

' Prueba de error: Sin entorno configurado lanza error
Private Function Test_CreateConfigService_SinEntornoConfigurado_LanzaError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_CreateConfigService_SinEntornoConfigurado_LanzaError"
    
    On Error GoTo ErrorHandler
    
    ' Arrange: Crear base de datos mock con TbLocalConfig vacía
    Dim db As DAO.Database
    Set db = CreateEmptyMockDatabase()
    
    ' Act: Intentar crear el servicio de configuración
    Dim errorOccurred As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    
    errorOccurred = False
    On Error Resume Next
    Dim config As IConfig
    Set config = modConfig.CreateConfigService(db)
    If Err.Number <> 0 Then
        errorOccurred = True
        errorNumber = Err.Number
        errorDescription = Err.Description
    End If
    On Error GoTo ErrorHandler
    
    ' Assert: Verificar que se lanzó el error esperado
    modAssert.AssertTrue errorOccurred, "Debe lanzarse un error cuando TbLocalConfig está vacía"
    modAssert.AssertEqual errorNumber, vbObjectError + 1001, "Debe lanzarse el error específico 1001"
    modAssert.AssertTrue InStr(errorDescription, "No se encontró configuración de entorno") > 0, "El mensaje de error debe mencionar la falta de configuración de entorno"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set Test_CreateConfigService_SinEntornoConfigurado_LanzaError = testResult
End Function

' Prueba de error: Entorno inválido lanza error
Private Function Test_CreateConfigService_ConEntornoInvalido_LanzaError() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_CreateConfigService_ConEntornoInvalido_LanzaError"
    
    On Error GoTo ErrorHandler
    
    ' Arrange: Crear base de datos mock con entorno inválido
    Dim db As DAO.Database
    Set db = CreateMockDatabaseWithEntorno("PRUEBAS")
    
    ' Act: Intentar crear el servicio de configuración
    Dim errorOccurred As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    
    errorOccurred = False
    On Error Resume Next
    Dim config As IConfig
    Set config = modConfig.CreateConfigService(db)
    If Err.Number <> 0 Then
        errorOccurred = True
        errorNumber = Err.Number
        errorDescription = Err.Description
    End If
    On Error GoTo ErrorHandler
    
    ' Assert: Verificar que se lanzó el error esperado
    modAssert.AssertTrue errorOccurred, "Debe lanzarse un error cuando el entorno es inválido"
    modAssert.AssertEqual errorNumber, vbObjectError + 1002, "Debe lanzarse el error específico 1002"
    modAssert.AssertTrue InStr(errorDescription, "Entorno no válido") > 0, "El mensaje de error debe mencionar entorno no válido"
    modAssert.AssertTrue InStr(errorDescription, "PRUEBAS") > 0, "El mensaje de error debe mencionar el valor inválido"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set Test_CreateConfigService_ConEntornoInvalido_LanzaError = testResult
End Function

'******************************************************************************
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK
'******************************************************************************

' Función auxiliar para crear una base de datos mock con un entorno específico
Private Function CreateMockDatabaseWithEntorno(ByVal entorno As String) As DAO.Database
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rs As DAO.recordset
    
    ' Crear base de datos temporal en memoria
    Set db = DBEngine.CreateDatabase("", dbLangGeneral, dbVersion120)
    
    ' Crear tabla TbLocalConfig
    Set tdf = db.CreateTableDef("TbLocalConfig")
    Set fld = tdf.CreateField("Entorno", dbText, 50)
    tdf.Fields.Append fld
    db.TableDefs.Append tdf
    
    ' Abrir recordset y añadir datos
    Set rs = db.OpenRecordset("TbLocalConfig", dbOpenDynaset)
    rs.AddNew
    rs.Fields("Entorno").value = entorno
    rs.Update
    rs.Close
    
    Set CreateMockDatabaseWithEntorno = db
End Function

' Función auxiliar para crear una base de datos mock vacía
Private Function CreateEmptyMockDatabase() As DAO.Database
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    ' Crear base de datos temporal en memoria
    Set db = DBEngine.CreateDatabase("", dbLangGeneral, dbVersion120)
    
    ' Crear tabla TbLocalConfig vacía
    Set tdf = db.CreateTableDef("TbLocalConfig")
    Set fld = tdf.CreateField("Entorno", dbText, 50)
    tdf.Fields.Append fld
    db.TableDefs.Append tdf
    
    Set CreateEmptyMockDatabase = db
End Function