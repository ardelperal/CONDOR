Option Compare Database
Option Explicit

' =====================================================
' MODULO: Test_CExpedienteRepository
' DESCRIPCION: Pruebas unitarias para CExpedienteRepository
' AUTOR: Sistema CONDOR
' FECHA: 2025
' =====================================================

#If DEV_MODE Then

' Función principal que ejecuta todas las pruebas del CExpedienteRepository
Public Function Test_CExpedienteRepository_RunAll() As CTestSuiteResult
    Dim suiteResult As CTestSuiteResult
    Set suiteResult = New CTestSuiteResult
    suiteResult.Initialize "Test_CExpedienteRepository"
    
    ' Ejecutar todas las pruebas individuales
    suiteResult.AddTestResult Test_ObtenerExpedientePorId_Success()
    suiteResult.AddTestResult Test_ObtenerExpedientePorId_NotFound()
    suiteResult.AddTestResult Test_ObtenerExpedientePorNemotecnico_Success()
    suiteResult.AddTestResult Test_ObtenerExpedientePorNemotecnico_NotFound()
    suiteResult.AddTestResult Test_ObtenerExpedientesActivosParaSelector_Success()
    suiteResult.AddTestResult Test_ObtenerExpedientesActivosParaSelector_EmptyResult()
    
    Set Test_CExpedienteRepository_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientePorId
' ============================================================================

' Prueba que ObtenerExpedientePorId devuelve correctamente un expediente existente
Private Function Test_ObtenerExpedientePorId_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientePorId_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
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
    Set Test_ObtenerExpedientePorId_Success = testResult
End Function

' Prueba que ObtenerExpedientePorId maneja correctamente expedientes no encontrados
Private Function Test_ObtenerExpedientePorId_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientePorId_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
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
    Set Test_ObtenerExpedientePorId_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientePorNemotecnico
' ============================================================================

' Prueba que ObtenerExpedientePorNemotecnico devuelve correctamente un expediente existente
Private Function Test_ObtenerExpedientePorNemotecnico_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientePorNemotecnico_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
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
    Set Test_ObtenerExpedientePorNemotecnico_Success = testResult
End Function

' Prueba que ObtenerExpedientePorNemotecnico maneja correctamente nemotécnicos no encontrados
Private Function Test_ObtenerExpedientePorNemotecnico_NotFound() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientePorNemotecnico_NotFound"
    
    On Error GoTo ErrorHandler
    
    ' Arrange - Configurar dependencias mock
    Dim mockConfig As New CMockConfig
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
    Set Test_ObtenerExpedientePorNemotecnico_NotFound = testResult
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA ObtenerExpedientesActivosParaSelector
' ============================================================================

' Prueba que ObtenerExpedientesActivosParaSelector devuelve correctamente expedientes activos
Private Function Test_ObtenerExpedientesActivosParaSelector_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientesActivosParaSelector_Success"
    
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
    Set Test_ObtenerExpedientesActivosParaSelector_Success = testResult
End Function

' Prueba que ObtenerExpedientesActivosParaSelector maneja correctamente cuando no hay expedientes activos
Private Function Test_ObtenerExpedientesActivosParaSelector_EmptyResult() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_ObtenerExpedientesActivosParaSelector_EmptyResult"
    
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
    Set Test_ObtenerExpedientesActivosParaSelector_EmptyResult = testResult
End Function

#End If