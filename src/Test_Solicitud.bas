Option Compare Database
Option Explicit
' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudService
' Arquitectura: Pruebas Aisladas con Inyección de Dependencias y Mocks
' Version: 2.0 - Refactorización Arquitectónica
' Fecha: 2025-01-14
' ============================================================================

' Función principal que ejecuta todas las pruebas de la suite
Public Function Test_Solicitud_RunAll() As CTestSuiteResult
    Dim result As CTestSuiteResult
    Set result = New CTestSuiteResult
    result.SuiteName = "Test_Solicitud"
    
    ' Ejecutar todas las pruebas unitarias
    result.AddTestResult Test_CreateSolicitud_Success()
    result.AddTestResult Test_SaveSolicitud_CallsRepository()
    
    Set Test_Solicitud_RunAll = result
End Function

' ============================================================================
' PRUEBAS UNITARIAS PARA CSolicitudService
' ============================================================================

' Prueba: CreateSolicitud debe crear correctamente una nueva solicitud
Public Function Test_CreateSolicitud_Success() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.TestName = "Test_CreateSolicitud_Success"
    
    On Error GoTo TestError
    
    ' ARRANGE: Configurar el servicio y los mocks
    Dim service As CSolicitudService
    Set service = New CSolicitudService
    
    ' Crear mocks
    Dim mockRepository As CMockSolicitudRepository
    Set mockRepository = New CMockSolicitudRepository
    
    Dim mockLogger As CMockOperationLogger
    Set mockLogger = New CMockOperationLogger
    
    ' Configurar el mock del repositorio para devolver un ID válido
    mockRepository.GuardarSolicitudReturnValue = 123
    
    ' Inyectar dependencias en el servicio
    service.Initialize mockRepository, mockLogger
    
    ' ACT: Ejecutar el método bajo prueba
    Dim resultado As T_Solicitud
    Set resultado = service.CreateSolicitud("EXP-001", "PC", "")
    
    ' ASSERT: Verificar los resultados
    
    ' 1. Verificar que el método GuardarSolicitud del mock fue llamado una vez
    If Not mockRepository.GuardarSolicitudCalled Then
        testResult.Success = False
        testResult.ErrorMessage = "El método GuardarSolicitud del repositorio no fue llamado"
        GoTo TestExit
    End If
    
    ' 2. Verificar que el objeto T_Solicitud tiene estadoInterno = "Borrador"
    If Not mockRepository.LastSavedSolicitud Is Nothing Then
        If mockRepository.LastSavedSolicitud.estadoInterno <> "Borrador" Then
            testResult.Success = False
            testResult.ErrorMessage = "El estadoInterno no es 'Borrador', es: " & mockRepository.LastSavedSolicitud.estadoInterno
            GoTo TestExit
        End If
    Else
        testResult.Success = False
        testResult.ErrorMessage = "No se guardó ninguna solicitud en el mock del repositorio"
        GoTo TestExit
    End If
    
    ' 3. Verificar que LogOperation del logger fue llamado
    If Not mockLogger.LogOperationCalled Then
        testResult.Success = False
        testResult.ErrorMessage = "El método LogOperation del logger no fue llamado"
        GoTo TestExit
    End If
    
    ' 4. Verificar que la función devolvió un objeto T_Solicitud no nulo
    If resultado Is Nothing Then
        testResult.Success = False
        testResult.ErrorMessage = "La función CreateSolicitud devolvió Nothing"
        GoTo TestExit
    End If
    
    ' 5. Verificar que el ID fue asignado correctamente
    If resultado.idSolicitud <> 123 Then
        testResult.Success = False
        testResult.ErrorMessage = "El ID de la solicitud no es el esperado. Esperado: 123, Actual: " & resultado.idSolicitud
        GoTo TestExit
    End If
    
    ' Si llegamos aquí, la prueba fue exitosa
    testResult.Success = True
    testResult.ErrorMessage = ""
    
TestExit:
    Set Test_CreateSolicitud_Success = testResult
    Exit Function
    
TestError:
    testResult.Success = False
    testResult.ErrorMessage = "Error en la prueba: " & Err.Description
    Set Test_CreateSolicitud_Success = testResult
End Function

' Prueba: SaveSolicitud debe llamar al repositorio con el objeto correcto
Public Function Test_SaveSolicitud_CallsRepository() As CTestResult
    Dim testResult As CTestResult
    Set testResult = New CTestResult
    testResult.TestName = "Test_SaveSolicitud_CallsRepository"
    
    On Error GoTo TestError
    
    ' ARRANGE: Configurar el servicio y los mocks
    Dim service As CSolicitudService
    Set service = New CSolicitudService
    
    ' Crear mocks
    Dim mockRepository As CMockSolicitudRepository
    Set mockRepository = New CMockSolicitudRepository
    
    Dim mockLogger As CMockOperationLogger
    Set mockLogger = New CMockOperationLogger
    
    ' Configurar el mock del repositorio
    mockRepository.GuardarSolicitudReturnValue = 456
    
    ' Inyectar dependencias
    service.Initialize mockRepository, mockLogger
    
    ' Crear un objeto T_Solicitud de prueba
    Dim miSolicitudDePrueba As T_Solicitud
    Set miSolicitudDePrueba = New T_Solicitud
    miSolicitudDePrueba.idSolicitud = 100
    miSolicitudDePrueba.idExpediente = "EXP-TEST"
    miSolicitudDePrueba.tipoSolicitud = "PC"
    miSolicitudDePrueba.estadoInterno = "EnProceso"
    
    ' ACT: Ejecutar el método bajo prueba
    Dim resultado As Boolean
    resultado = service.SaveSolicitud(miSolicitudDePrueba)
    
    ' ASSERT: Verificar los resultados
    
    ' 1. Verificar que el método GuardarSolicitud del mock fue llamado
    If Not mockRepository.GuardarSolicitudCalled Then
        testResult.Success = False
        testResult.ErrorMessage = "El método GuardarSolicitud del repositorio no fue llamado"
        GoTo TestExit
    End If
    
    ' 2. Verificar que se llamó con el objeto T_Solicitud exacto
    If mockRepository.LastSavedSolicitud Is Nothing Then
        testResult.Success = False
        testResult.ErrorMessage = "No se guardó ninguna solicitud en el mock del repositorio"
        GoTo TestExit
    End If
    
    ' Verificar que es el mismo objeto (comparar propiedades clave)
    If mockRepository.LastSavedSolicitud.idSolicitud <> miSolicitudDePrueba.idSolicitud Then
        testResult.Success = False
        testResult.ErrorMessage = "El ID de la solicitud guardada no coincide. Esperado: " & miSolicitudDePrueba.idSolicitud & ", Actual: " & mockRepository.LastSavedSolicitud.idSolicitud
        GoTo TestExit
    End If
    
    If mockRepository.LastSavedSolicitud.idExpediente <> miSolicitudDePrueba.idExpediente Then
        testResult.Success = False
        testResult.ErrorMessage = "El idExpediente de la solicitud guardada no coincide"
        GoTo TestExit
    End If
    
    ' 3. Verificar que el resultado es True (porque el mock devuelve un ID > 0)
    If Not resultado Then
        testResult.Success = False
        testResult.ErrorMessage = "SaveSolicitud debería devolver True cuando el repositorio devuelve un ID > 0"
        GoTo TestExit
    End If
    
    ' Si llegamos aquí, la prueba fue exitosa
    testResult.Success = True
    testResult.ErrorMessage = ""
    
TestExit:
    Set Test_SaveSolicitud_CallsRepository = testResult
    Exit Function
    
TestError:
    testResult.Success = False
    testResult.ErrorMessage = "Error en la prueba: " & Err.Description
    Set Test_SaveSolicitud_CallsRepository = testResult
End Function