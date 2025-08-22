Attribute VB_Name = "Test_CConfig"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CConfig
' ============================================================================
' Este módulo contiene pruebas unitarias aisladas para CConfig
' utilizando mocks para todas las dependencias externas.
' Sigue la Lección 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable

' Función principal que ejecuta todas las pruebas del módulo
Public Function Test_CConfig_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "Test_CConfig"
    
    ' Ejecutar todas las pruebas unitarias
    suiteResult.AddTestResult Test_Initialize_LoadsAndReadsValues_Success()
    
    Set Test_CConfig_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA CConfig
' ============================================================================

' Prueba que CConfig inicializa correctamente y lee valores desde el mock repository
Private Function Test_Initialize_LoadsAndReadsValues_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "Test_Initialize_LoadsAndReadsValues_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar mocks y datos de prueba
    Dim mockRepository As New CMockSolicitudRepository
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear recordset mock con claves de configuración
    Dim mockRecordset As Object
    Set mockRecordset = CreateMockConfigRecordset()
    mockRepository.SetMockRecordset mockRecordset
    
    ' Crear instancia de CConfig con dependencias mock
    Dim config As New CConfig
    config.Initialize mockLogger, mockRepository
    
    ' Act - Ejecutar inicialización y métodos bajo prueba
    Dim initResult As Boolean
    initResult = config.InitializeEnvironment()
    
    ' Assert - Verificar que la inicialización fue exitosa
    modAssert.AssertTrue initResult, "InitializeEnvironment debe retornar True"
    
    ' Verificar que GetValue devuelve los valores correctos del recordset mock
    modAssert.AssertEquals "C:\\Test\\Database.accdb", config.GetValue("DATABASEPATH"), "DATABASEPATH debe coincidir con el valor mock"
    modAssert.AssertEquals "C:\\Test\\Logs", config.GetValue("LOGPATH"), "LOGPATH debe coincidir con el valor mock"
    
    ' Verificar que HasKey funciona correctamente para claves existentes e inexistentes
    modAssert.AssertTrue config.HasKey("DATABASEPATH"), "HasKey debe retornar True para claves existentes"
    modAssert.AssertTrue config.HasKey("LOGPATH"), "HasKey debe retornar True para claves existentes"
    modAssert.AssertFalse config.HasKey("CLAVE_INEXISTENTE"), "HasKey debe retornar False para claves inexistentes"
    
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
    Set Test_Initialize_LoadsAndReadsValues_Success = testResult
End Function

' ============================================================================
' FUNCIONES AUXILIARES PARA CREAR RECORDSETS MOCK
' ============================================================================

' Crea un recordset mock con datos de configuración para las pruebas
Private Function CreateMockConfigRecordset() As Object
    Dim rs As Object
    
    ' Crear recordset ADODB en memoria
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Definir campos del recordset
    rs.Fields.Append "Clave", 202, 50 ' adVarWChar
    rs.Fields.Append "Valor", 202, 255 ' adVarWChar
    
    ' Abrir recordset en memoria
    rs.Open
    
    ' Añadir registro DATABASEPATH
    rs.AddNew
    rs.Fields("Clave").Value = "DATABASEPATH"
    rs.Fields("Valor").Value = "C:\\Test\\Database.accdb"
    rs.Update
    
    ' Añadir registro LOGPATH
    rs.AddNew
    rs.Fields("Clave").Value = "LOGPATH"
    rs.Fields("Valor").Value = "C:\\Test\\Logs"
    rs.Update
    
    ' Mover al primer registro
    rs.MoveFirst
    
    Set CreateMockConfigRecordset = rs
End Function







