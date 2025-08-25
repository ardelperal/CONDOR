Attribute VB_Name = "IntegrationTest_CMapeoRepository"
Option Compare Database
Option Explicit

' =====================================================
' MODULO: IntegrationTest_CMapeoRepository
' DESCRIPCION: Pruebas de integración para CMapeoRepository
' AUTOR: Sistema CONDOR
' FECHA: 2025
' =====================================================

#If DEV_MODE Then

' Función principal que ejecuta todas las pruebas de integración del CMapeoRepository
Public Function IntegrationTest_CMapeoRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_CMapeoRepository"
    
    ' Ejecutar todas las pruebas individuales
    suiteResult.AddTestResult IntegrationTest_GetMapeoPorTipo_Success()
    suiteResult.AddTestResult IntegrationTest_GetMapeoPorTipo_NotFound()
    suiteResult.AddTestResult IntegrationTest_GetMapeoPorTipo_EmptyType()
    
    Set IntegrationTest_CMapeoRepository_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA GetMapeoPorTipo
' ============================================================================

' Prueba que GetMapeoPorTipo devuelve correctamente un mapeo existente
Private Function IntegrationTest_GetMapeoPorTipo_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_GetMapeoPorTipo_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CMapeoRepository
    repository.Initialize mockConfig
    
    ' Act - Ejecutar el método bajo prueba con un tipo conocido
    Dim rs As DAO.Recordset
    Set rs = repository.GetMapeoPorTipo("PC")
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "GetMapeoPorTipo debe devolver un recordset válido"
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set IntegrationTest_GetMapeoPorTipo_Success = testResult
End Function

' Prueba que GetMapeoPorTipo maneja correctamente tipos no encontrados
Private Function IntegrationTest_GetMapeoPorTipo_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_GetMapeoPorTipo_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CMapeoRepository
    repository.Initialize mockConfig
    
    ' Act - Ejecutar el método con tipo inexistente
    Dim rs As DAO.Recordset
    Set rs = repository.GetMapeoPorTipo("TIPO_INEXISTENTE")
    
    ' Assert - Verificar que maneja correctamente el caso no encontrado
    If Not rs Is Nothing Then
        modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para tipo no encontrado"
    End If
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set IntegrationTest_GetMapeoPorTipo_NotFound = testResult
End Function

' Prueba que GetMapeoPorTipo maneja correctamente tipos vacíos
Private Function IntegrationTest_GetMapeoPorTipo_EmptyType() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_GetMapeoPorTipo_EmptyType"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CMapeoRepository
    repository.Initialize mockConfig
    
    ' Act - Ejecutar el método con tipo vacío
    Dim rs As DAO.Recordset
    Set rs = repository.GetMapeoPorTipo("")
    
    ' Assert - Verificar que maneja correctamente el caso de tipo vacío
    If Not rs Is Nothing Then
        modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para tipo vacío"
    End If
    
    testResult.Pass
    GoTo CleanUp
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Description
    
CleanUp:
    ' Limpiar recursos
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set IntegrationTest_GetMapeoPorTipo_EmptyType = testResult
End Function

#End If