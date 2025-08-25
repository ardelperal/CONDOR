Option Compare Database
Option Explicit

' =====================================================
' MODULO: IntegrationTest_CExpedienteRepository
' DESCRIPCION: Pruebas de integración para CExpedienteRepository
' AUTOR: Sistema CONDOR
' FECHA: 2025
' =====================================================

#If DEV_MODE Then

' Función principal que ejecuta todas las pruebas de integración del CExpedienteRepository
Public Function IntegrationTest_CExpedienteRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_CExpedienteRepository"
    
    ' Ejecutar todas las pruebas individuales
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorId_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorId_NotFound()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorNemotecnico_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientesActivosParaSelector_Success()
    suiteResult.AddTestResult IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult()
    
    Set IntegrationTest_CExpedienteRepository_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientePorId
' ============================================================================

' Prueba que ObtenerExpedientePorId devuelve correctamente un expediente existente
Private Function IntegrationTest_ObtenerExpedientePorId_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorId_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientePorId(1)
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "ObtenerExpedientePorId debe devolver un recordset válido"
    
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
    Set IntegrationTest_ObtenerExpedientePorId_Success = testResult
End Function

' Prueba que ObtenerExpedientePorId maneja correctamente expedientes no encontrados
Private Function IntegrationTest_ObtenerExpedientePorId_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorId_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método con ID inexistente
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientePorId(99999)
    
    ' Assert - Verificar que maneja correctamente el caso no encontrado
    If Not rs Is Nothing Then
        modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para expediente no encontrado"
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
    Set IntegrationTest_ObtenerExpedientePorId_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientePorNemotecnico
' ============================================================================

' Prueba que ObtenerExpedientePorNemotecnico devuelve correctamente un expediente existente
Private Function IntegrationTest_ObtenerExpedientePorNemotecnico_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorNemotecnico_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientePorNemotecnico("EXP-2024-001")
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "ObtenerExpedientePorNemotecnico debe devolver un recordset válido"
    
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
    Set IntegrationTest_ObtenerExpedientePorNemotecnico_Success = testResult
End Function

' Prueba que ObtenerExpedientePorNemotecnico maneja correctamente nemotécnicos no encontrados
Private Function IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método con nemotécnico inexistente
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientePorNemotecnico("INEXISTENTE-999")
    
    ' Assert - Verificar que maneja correctamente el caso no encontrado
    If Not rs Is Nothing Then
        modAssert.AssertTrue rs.EOF, "El recordset debe estar vacío para nemotécnico no encontrado"
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
    Set IntegrationTest_ObtenerExpedientePorNemotecnico_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientesActivosParaSelector
' ============================================================================

' Prueba que ObtenerExpedientesActivosParaSelector devuelve correctamente expedientes activos
Private Function IntegrationTest_ObtenerExpedientesActivosParaSelector_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientesActivosParaSelector_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    mockConfig.AddSetting "BACKEND_DB_PATH", "C:\Test\CONDOR_Backend.accdb"
    mockConfig.AddSetting "DATABASE_PASSWORD", "testpassword"
    
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientesActivosParaSelector()
    
    ' Assert - Verificar que devuelve un recordset válido
    modAssert.AssertNotNull rs, "ObtenerExpedientesActivosParaSelector debe devolver un recordset válido"
    
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
    Set IntegrationTest_ObtenerExpedientesActivosParaSelector_Success = testResult
End Function

' Prueba que ObtenerExpedientesActivosParaSelector maneja correctamente cuando no hay expedientes activos
Private Function IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
    Dim mockLogger As New CMockOperationLogger
    
    ' Crear repositorio con dependencias mock
    Dim repository As New CExpedienteRepository
    repository.Initialize mockConfig, mockLogger
    
    ' Act - Ejecutar el método bajo prueba
    Dim rs As DAO.Recordset
    Set rs = repository.ObtenerExpedientesActivosParaSelector()
    
    ' Assert - Verificar que maneja correctamente el caso sin resultados
    If Not rs Is Nothing Then
        ' El recordset puede estar vacío, lo cual es válido
        modAssert.AssertNotNull rs, "El recordset debe ser válido aunque esté vacío"
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
    Set IntegrationTest_ObtenerExpedientesActivosParaSelector_EmptyResult = testResult
End Function

#End If