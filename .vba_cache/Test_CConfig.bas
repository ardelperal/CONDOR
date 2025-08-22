Option Compare Database
Option Explicit
' Módulo: Test_CConfig
' Descripción: Pruebas de integración para CConfig usando framework de testing
' Arquitectura: Capa de Pruebas - Tests de Integración

' Test de integración principal: Verificar que CConfig carga configuraciones desde la base de datos
Public Function Test_Initialize_LoadsSettingsFromDatabase_Success() As CTestResult
    On Error GoTo ErrorHandler
    
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.TestName = "Test_Initialize_LoadsSettingsFromDatabase_Success"
    
    ' Crear mock repository que devuelve un recordset falso
    Dim mockRepository As CMockSolicitudRepository
    Set mockRepository = New CMockSolicitudRepository
    
    ' Configurar el mock para devolver datos de configuración
    Dim mockRecordset As Object
    Set mockRecordset = CreateMockRecordset()
    mockRepository.SetMockConfigurationSettings mockRecordset
    
    ' Crear instancia real de CConfig
    Dim config As CConfig
    Set config = New CConfig
    
    ' Crear mock operation logger
    Dim mockLogger As CMockOperationLogger
    Set mockLogger = New CMockOperationLogger
    
    ' Inyectar dependencias
    config.Initialize mockLogger, mockRepository
    
    ' Llamar a InitializeEnvironment
    Dim initResult As Boolean
    initResult = config.InitializeEnvironment()
    
    ' Verificar que la inicialización fue exitosa
    Call modAssert.AssertTrue(initResult, "InitializeEnvironment debería retornar True")
    
    ' Verificar que los valores se cargaron correctamente
    Call modAssert.AssertEqual("C:\\Test\\Database.accdb", config.GetValue("DATABASEPATH"), "DatabasePath debería cargarse correctamente")
    Call modAssert.AssertEqual("C:\\Test\\Data.accdb", config.GetValue("DATAPATH"), "DataPath debería cargarse correctamente")
    Call modAssert.AssertEqual("C:\\Test\\Logs", config.GetValue("LOGPATH"), "LogPath debería cargarse correctamente")
    
    ' Verificar que HasKey funciona correctamente
    Call modAssert.AssertTrue(config.HasKey("DATABASEPATH"), "HasKey debería retornar True para claves existentes")
    Call modAssert.AssertFalse(config.HasKey("CLAVE_INEXISTENTE"), "HasKey debería retornar False para claves inexistentes")
    
    testResult.Success = True
    testResult.Message = "Test completado exitosamente"
    Set Test_Initialize_LoadsSettingsFromDatabase_Success = testResult
    
    Exit Function
    
ErrorHandler:
    testResult.Success = False
    testResult.Message = "Error en test: " & Err.Description
    Set Test_Initialize_LoadsSettingsFromDatabase_Success = testResult
End Function

' Función auxiliar para crear un recordset mock con datos de configuración
Private Function CreateMockRecordset() As Object
    On Error GoTo ErrorHandler
    
    ' Crear recordset mock usando ADODB
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos
    rs.Fields.Append "Clave", 200, 50  ' adVarChar
    rs.Fields.Append "Valor", 200, 255 ' adVarChar
    rs.Open
    
    ' Agregar datos de prueba
    rs.AddNew
    rs.Fields("Clave").value = "DatabasePath"
    rs.Fields("Valor").value = "C:\\Test\\Database.accdb"
    rs.Update
    
    rs.AddNew
    rs.Fields("Clave").value = "DataPath"
    rs.Fields("Valor").value = "C:\\Test\\Data.accdb"
    rs.Update
    
    rs.AddNew
    rs.Fields("Clave").value = "LogPath"
    rs.Fields("Valor").value = "C:\\Test\\Logs"
    rs.Update
    
    ' Mover al primer registro
    rs.MoveFirst
    
    Set CreateMockRecordset = rs
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "Test_CConfig.CreateMockRecordset"
    Set CreateMockRecordset = Nothing
End Function

' Función principal para ejecutar todas las pruebas de CConfig
Public Function RunAllCConfigTests() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.SuiteName = "Test_CConfig Suite"
    
    ' Ejecutar test de integración
    Dim testResult As CTestResult
    Set testResult = Test_Initialize_LoadsSettingsFromDatabase_Success()
    suiteResult.AddTestResult testResult
    
    Set RunAllCConfigTests = suiteResult
    Exit Function
    
ErrorHandler:
    modErrorHandler.LogError Err.Number, Err.Description, "Test_CConfig.RunAllCConfigTests"
    Set RunAllCConfigTests = Nothing
End Function








