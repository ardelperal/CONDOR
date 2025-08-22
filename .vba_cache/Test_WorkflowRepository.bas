Attribute VB_Name = "Test_WorkflowRepository"
'==============================================================================
' Módulo: Test_WorkflowRepository
' Propósito: Pruebas de integración para CWorkflowRepository
' Autor: CONDOR-Expert
' Fecha: 2024
'==============================================================================

Option Compare Database
Option Explicit

'==============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'==============================================================================

'''
' Ejecuta todas las pruebas de integración del WorkflowRepository
' @return CTestSuiteResult: Resultado de la suite de pruebas
'''
Public Function Test_WorkflowRepository_RunAll() As CTestSuiteResult
    On Error GoTo ErrorHandler
    
    ' Crear la suite de resultados
    Dim suite As CTestSuiteResult
    Set suite = New CTestSuiteResult
    suite.Initialize "Test_WorkflowRepository", "Pruebas de integración para CWorkflowRepository"
    
    ' Ejecutar todas las pruebas individuales
    suite.AddTestResult Test_WorkflowRepository_ValidTransition_ReturnsTrue()
    suite.AddTestResult Test_WorkflowRepository_InvalidTransition_ReturnsFalse()
    suite.AddTestResult Test_WorkflowRepository_NonExistentType_ReturnsFalse()
    suite.AddTestResult Test_WorkflowRepository_InactiveTransition_ReturnsFalse()
    
    Set Test_WorkflowRepository_RunAll = suite
    Exit Function
    
ErrorHandler:
    If suite Is Nothing Then Set suite = New CTestSuiteResult
    suite.Initialize "Test_WorkflowRepository", "Error en suite de pruebas"
    Set Test_WorkflowRepository_RunAll = suite
End Function

'==============================================================================
' PRUEBAS DE INTEGRACIÓN
'==============================================================================

'''
' Prueba que una transición válida existente en la base de datos devuelve True
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_ValidTransition_ReturnsTrue() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_ValidTransition_ReturnsTrue", "Transición válida debe devolver True"
    
    On Error GoTo ErrorHandler
    
    ' Preparar datos de prueba en la base de datos
    Call SetupTestData_ValidTransition
    
    ' Crear instancia real del repositorio usando factory
    Dim repository As IWorkflowRepository
    Set repository = modWorkflowRepositoryFactory.CreateWorkflowRepository()
    
    ' Ejecutar la prueba
    Dim result As Boolean
    result = repository.IsValidTransition("PC", "BORRADOR", "ENVIADO")
    
    ' Verificar resultado
    If result Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba True para transición válida PC: BORRADOR -> ENVIADO"
    End If
    
    ' Limpiar datos de prueba
    Call CleanupTestData
    
    Set Test_WorkflowRepository_ValidTransition_ReturnsTrue = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Call CleanupTestData
    Set Test_WorkflowRepository_ValidTransition_ReturnsTrue = testResult
End Function

'''
' Prueba que una transición inválida devuelve False
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_InvalidTransition_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_InvalidTransition_ReturnsFalse", "Transición inválida debe devolver False"
    
    On Error GoTo ErrorHandler
    
    ' Preparar datos de prueba en la base de datos
    Call SetupTestData_ValidTransition
    
    ' Crear instancia real del repositorio usando factory
    Dim repository As IWorkflowRepository
    Set repository = modWorkflowRepositoryFactory.CreateWorkflowRepository()
    
    ' Ejecutar la prueba con una transición que no existe
    Dim result As Boolean
    result = repository.IsValidTransition("PC", "ENVIADO", "BORRADOR")
    
    ' Verificar resultado
    If Not result Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False para transición inválida PC: ENVIADO -> BORRADOR"
    End If
    
    ' Limpiar datos de prueba
    Call CleanupTestData
    
    Set Test_WorkflowRepository_InvalidTransition_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Call CleanupTestData
    Set Test_WorkflowRepository_InvalidTransition_ReturnsFalse = testResult
End Function

'''
' Prueba que un tipo de solicitud inexistente devuelve False
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_NonExistentType_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_NonExistentType_ReturnsFalse", "Tipo inexistente debe devolver False"
    
    On Error GoTo ErrorHandler
    
    ' Preparar datos de prueba en la base de datos
    Call SetupTestData_ValidTransition
    
    ' Crear instancia real del repositorio usando factory
    Dim repository As IWorkflowRepository
    Set repository = modWorkflowRepositoryFactory.CreateWorkflowRepository()
    
    ' Ejecutar la prueba con un tipo inexistente
    Dim result As Boolean
    result = repository.IsValidTransition("INEXISTENTE", "BORRADOR", "ENVIADO")
    
    ' Verificar resultado
    If Not result Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False para tipo de solicitud inexistente"
    End If
    
    ' Limpiar datos de prueba
    Call CleanupTestData
    
    Set Test_WorkflowRepository_NonExistentType_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Call CleanupTestData
    Set Test_WorkflowRepository_NonExistentType_ReturnsFalse = testResult
End Function

'''
' Prueba que una transición inactiva devuelve False
' @return CTestResult: Resultado de la prueba
'''
Private Function Test_WorkflowRepository_InactiveTransition_ReturnsFalse() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.Initialize "Test_WorkflowRepository_InactiveTransition_ReturnsFalse", "Transición inactiva debe devolver False"
    
    On Error GoTo ErrorHandler
    
    ' Preparar datos de prueba con transición inactiva
    Call SetupTestData_InactiveTransition
    
    ' Crear instancia real del repositorio usando factory
    Dim repository As IWorkflowRepository
    Set repository = modWorkflowRepositoryFactory.CreateWorkflowRepository()
    
    ' Ejecutar la prueba
    Dim result As Boolean
    result = repository.IsValidTransition("PC", "BORRADOR", "CANCELADO")
    
    ' Verificar resultado
    If Not result Then
        testResult.Pass
    Else
        testResult.Fail "Se esperaba False para transición inactiva PC: BORRADOR -> CANCELADO"
    End If
    
    ' Limpiar datos de prueba
    Call CleanupTestData
    
    Set Test_WorkflowRepository_InactiveTransition_ReturnsFalse = testResult
    Exit Function
    
ErrorHandler:
    testResult.Fail "Error durante la prueba: " & Err.Number & " - " & Err.Description
    Call CleanupTestData
    Set Test_WorkflowRepository_InactiveTransition_ReturnsFalse = testResult
End Function

'==============================================================================
' FUNCIONES AUXILIARES PARA PREPARACIÓN DE DATOS
'==============================================================================

'''
' Configura datos de prueba para transiciones válidas
'''
Private Sub SetupTestData_ValidTransition()
    On Error GoTo ErrorHandler
    
    Dim configService As IConfig
    Set configService = modConfig.GetInstance()
    
    Dim backendPath As String
    backendPath = configService.GetValue("DATAPATH")
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(backendPath)
    
    ' Limpiar datos previos
    Call CleanupTestData
    
    ' Insertar estados de prueba
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado, NombreEstado) VALUES (9001, 'BORRADOR', 'Test Borrador')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado, NombreEstado) VALUES (9002, 'ENVIADO', 'Test Enviado')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado, NombreEstado) VALUES (9003, 'CANCELADO', 'Test Cancelado')"
    
    ' Insertar transición válida activa
    db.Execute "INSERT INTO TbTransiciones (ID, TipoSolicitud, EstadoOrigenID, EstadoDestinoID, Activo) " & _
               "VALUES (9001, 'PC', 9001, 9002, True)"
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "Error en SetupTestData_ValidTransition: " & Err.Number & " - " & Err.Description
End Sub

'''
' Configura datos de prueba para transiciones inactivas
'''
Private Sub SetupTestData_InactiveTransition()
    On Error GoTo ErrorHandler
    
    Dim configService As IConfig
    Set configService = modConfig.GetInstance()
    
    Dim backendPath As String
    backendPath = configService.GetValue("DATAPATH")
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(backendPath)
    
    ' Limpiar datos previos
    Call CleanupTestData
    
    ' Insertar estados de prueba
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado, NombreEstado) VALUES (9001, 'BORRADOR', 'Test Borrador')"
    db.Execute "INSERT INTO TbEstados (ID, CodigoEstado, NombreEstado) VALUES (9003, 'CANCELADO', 'Test Cancelado')"
    
    ' Insertar transición inactiva
    db.Execute "INSERT INTO TbTransiciones (ID, TipoSolicitud, EstadoOrigenID, EstadoDestinoID, Activo) " & _
               "VALUES (9002, 'PC', 9001, 9003, False)"
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "Error en SetupTestData_InactiveTransition: " & Err.Number & " - " & Err.Description
End Sub

'''
' Limpia todos los datos de prueba de las tablas
'''
Private Sub CleanupTestData()
    On Error GoTo ErrorHandler
    
    Dim configService As IConfig
    Set configService = modConfig.GetInstance()
    
    Dim backendPath As String
    backendPath = configService.GetValue("DATAPATH")
    
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(backendPath)
    
    ' Eliminar transiciones de prueba
    db.Execute "DELETE FROM TbTransiciones WHERE ID >= 9000"
    
    ' Eliminar estados de prueba
    db.Execute "DELETE FROM TbEstados WHERE ID >= 9000"
    
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not db Is Nothing Then Set db = Nothing
    Debug.Print "Error en CleanupTestData: " & Err.Number & " - " & Err.Description
End Sub