Attribute VB_Name = "Test_WordManager"
'******************************************************************************
' Módulo: Test_WordManager
' Descripción: Suite de pruebas unitarias puras para WordManager usando mocks.
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
' Versión: 2.0
' Nota: Cumple con Lección 10 - Pruebas unitarias puras sin dependencias externas
'******************************************************************************

Option Compare Database
Option Explicit

' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'******************************************************************************

Public Function Test_WordManager_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "Test_WordManager"
    
    ' Ejecutar todas las pruebas unitarias puras
    suite.AddTest Test_AbrirDocumento_Success
    suite.AddTest Test_AbrirDocumento_Failure
    suite.AddTest Test_ReemplazarTexto_Success
    suite.AddTest Test_ReemplazarTexto_Failure
    suite.AddTest Test_GuardarDocumento_Success
    suite.AddTest Test_GuardarDocumento_Failure
    suite.AddTest Test_LeerContenidoDocumento_Success
    suite.AddTest Test_LeerContenidoDocumento_Failure
    suite.AddTest Test_CerrarDocumento_Success
    
    Set Test_WordManager_RunAll = suite
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PURAS - ABRIR DOCUMENTO
'******************************************************************************

' Prueba unitaria: AbrirDocumento retorna True cuando el mock está configurado para éxito
Private Function Test_AbrirDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_AbrirDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    rutaDocumento = "C:\Test\documento.docx"
    
    ' Configurar mock para retornar éxito
    mockWordManager.AbrirDocumento_ReturnValue = True
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.AbrirDocumento(rutaDocumento)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "AbrirDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDocumento, mockWordManager.AbrirDocumento_LastRutaDocumento, "Debería pasar la ruta correcta"
    modAssert.AssertTrue resultado, "Debería retornar True cuando el mock está configurado para éxito"
    modAssert.AssertEqual 1, mockWordManager.AbrirDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: AbrirDocumento Success"
    
TestExit:
    Set Test_AbrirDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba unitaria: AbrirDocumento retorna False cuando el mock está configurado para fallo
Private Function Test_AbrirDocumento_Failure() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_AbrirDocumento_Failure"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    rutaDocumento = "C:\Test\archivo_inexistente.docx"
    
    ' Configurar mock para retornar fallo
    mockWordManager.AbrirDocumento_ReturnValue = False
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.AbrirDocumento(rutaDocumento)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "AbrirDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDocumento, mockWordManager.AbrirDocumento_LastRutaDocumento, "Debería pasar la ruta correcta"
    modAssert.AssertFalse resultado, "Debería retornar False cuando el mock está configurado para fallo"
    modAssert.AssertEqual 1, mockWordManager.AbrirDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: AbrirDocumento Failure"
    
TestExit:
    Set Test_AbrirDocumento_Failure = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PURAS - REEMPLAZAR TEXTO
'******************************************************************************

' Prueba unitaria: ReemplazarTexto retorna True cuando el mock está configurado para éxito
Private Function Test_ReemplazarTexto_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_ReemplazarTexto_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim textoABuscar As String
    Dim textoReemplazo As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    textoABuscar = "[MARCADOR]"
    textoReemplazo = "TEXTO_NUEVO"
    
    ' Configurar mock para retornar éxito
    mockWordManager.ReemplazarTexto_ReturnValue = True
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.ReemplazarTexto(textoABuscar, textoReemplazo)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.ReemplazarTexto_WasCalled, "ReemplazarTexto debería haber sido llamado"
    modAssert.AssertEqual textoABuscar, mockWordManager.ReemplazarTexto_LastTextoABuscar, "Debería pasar el texto a buscar correcto"
    modAssert.AssertEqual textoReemplazo, mockWordManager.ReemplazarTexto_LastTextoReemplazo, "Debería pasar el texto de reemplazo correcto"
    modAssert.AssertTrue resultado, "Debería retornar True cuando el mock está configurado para éxito"
    modAssert.AssertEqual 1, mockWordManager.ReemplazarTexto_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: ReemplazarTexto Success"
    
TestExit:
    Set Test_ReemplazarTexto_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba unitaria: ReemplazarTexto retorna False cuando el mock está configurado para fallo
Private Function Test_ReemplazarTexto_Failure() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_ReemplazarTexto_Failure"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim textoABuscar As String
    Dim textoReemplazo As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    textoABuscar = "[MARCADOR_INEXISTENTE]"
    textoReemplazo = "TEXTO_NUEVO"
    
    ' Configurar mock para retornar fallo
    mockWordManager.ReemplazarTexto_ReturnValue = False
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.ReemplazarTexto(textoABuscar, textoReemplazo)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.ReemplazarTexto_WasCalled, "ReemplazarTexto debería haber sido llamado"
    modAssert.AssertEqual textoABuscar, mockWordManager.ReemplazarTexto_LastTextoABuscar, "Debería pasar el texto a buscar correcto"
    modAssert.AssertEqual textoReemplazo, mockWordManager.ReemplazarTexto_LastTextoReemplazo, "Debería pasar el texto de reemplazo correcto"
    modAssert.AssertFalse resultado, "Debería retornar False cuando el mock está configurado para fallo"
    modAssert.AssertEqual 1, mockWordManager.ReemplazarTexto_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: ReemplazarTexto Failure"
    
TestExit:
    Set Test_ReemplazarTexto_Failure = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PURAS - GUARDAR DOCUMENTO
'******************************************************************************

' Prueba unitaria: GuardarDocumento retorna True cuando el mock está configurado para éxito
Private Function Test_GuardarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_GuardarDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDestino As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    rutaDestino = "C:\Test\documento_guardado.docx"
    
    ' Configurar mock para retornar éxito
    mockWordManager.GuardarDocumento_ReturnValue = True
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.GuardarDocumento(rutaDestino)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.GuardarDocumento_WasCalled, "GuardarDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDestino, mockWordManager.GuardarDocumento_LastRutaDestino, "Debería pasar la ruta de destino correcta"
    modAssert.AssertTrue resultado, "Debería retornar True cuando el mock está configurado para éxito"
    modAssert.AssertEqual 1, mockWordManager.GuardarDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: GuardarDocumento Success"
    
TestExit:
    Set Test_GuardarDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba unitaria: GuardarDocumento retorna False cuando el mock está configurado para fallo
Private Function Test_GuardarDocumento_Failure() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_GuardarDocumento_Failure"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDestino As String
    Dim resultado As Boolean
    
    ' Configurar datos de prueba
    rutaDestino = "C:\RutaInvalida\documento.docx"
    
    ' Configurar mock para retornar fallo
    mockWordManager.GuardarDocumento_ReturnValue = False
    
    ' Ejecutar método bajo prueba
    resultado = mockWordManager.GuardarDocumento(rutaDestino)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.GuardarDocumento_WasCalled, "GuardarDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDestino, mockWordManager.GuardarDocumento_LastRutaDestino, "Debería pasar la ruta de destino correcta"
    modAssert.AssertFalse resultado, "Debería retornar False cuando el mock está configurado para fallo"
    modAssert.AssertEqual 1, mockWordManager.GuardarDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: GuardarDocumento Failure"
    
TestExit:
    Set Test_GuardarDocumento_Failure = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PURAS - LEER CONTENIDO DOCUMENTO
'******************************************************************************

' Prueba unitaria: LeerContenidoDocumento retorna contenido cuando el mock está configurado
Private Function Test_LeerContenidoDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_LeerContenidoDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    Dim contenidoEsperado As String
    Dim contenidoObtenido As String
    
    ' Configurar datos de prueba
    rutaDocumento = "C:\Test\documento_lectura.docx"
    contenidoEsperado = "Contenido de prueba para validar lectura"
    
    ' Configurar mock para retornar contenido específico
    mockWordManager.LeerContenidoDocumento_ReturnValue = contenidoEsperado
    
    ' Ejecutar método bajo prueba
    contenidoObtenido = mockWordManager.LeerContenidoDocumento(rutaDocumento)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.LeerContenidoDocumento_WasCalled, "LeerContenidoDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDocumento, mockWordManager.LeerContenidoDocumento_LastRutaDocumento, "Debería pasar la ruta correcta"
    modAssert.AssertEqual contenidoEsperado, contenidoObtenido, "Debería retornar el contenido configurado en el mock"
    modAssert.AssertEqual 1, mockWordManager.LeerContenidoDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: LeerContenidoDocumento Success"
    
TestExit:
    Set Test_LeerContenidoDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba unitaria: LeerContenidoDocumento retorna cadena vacía cuando el mock está configurado para fallo
Private Function Test_LeerContenidoDocumento_Failure() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_LeerContenidoDocumento_Failure"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    Dim contenidoObtenido As String
    
    ' Configurar datos de prueba
    rutaDocumento = "C:\Test\documento_inexistente.docx"
    
    ' Configurar mock para retornar cadena vacía (simulando fallo)
    mockWordManager.LeerContenidoDocumento_ReturnValue = ""
    
    ' Ejecutar método bajo prueba
    contenidoObtenido = mockWordManager.LeerContenidoDocumento(rutaDocumento)
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.LeerContenidoDocumento_WasCalled, "LeerContenidoDocumento debería haber sido llamado"
    modAssert.AssertEqual rutaDocumento, mockWordManager.LeerContenidoDocumento_LastRutaDocumento, "Debería pasar la ruta correcta"
    modAssert.AssertEqual "", contenidoObtenido, "Debería retornar cadena vacía cuando el mock está configurado para fallo"
    modAssert.AssertEqual 1, mockWordManager.LeerContenidoDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: LeerContenidoDocumento Failure"
    
TestExit:
    Set Test_LeerContenidoDocumento_Failure = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

'******************************************************************************
' PRUEBAS UNITARIAS PURAS - CERRAR DOCUMENTO
'******************************************************************************

' Prueba unitaria: CerrarDocumento ejecuta correctamente
Private Function Test_CerrarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_CerrarDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    
    ' Ejecutar método bajo prueba
    mockWordManager.CerrarDocumento
    
    ' Aserciones sobre el mock
    modAssert.AssertTrue mockWordManager.CerrarDocumento_WasCalled, "CerrarDocumento debería haber sido llamado"
    modAssert.AssertEqual 1, mockWordManager.CerrarDocumento_CallCount, "Debería haber sido llamado exactamente una vez"
    
    testResult.Passed = True
    testResult.Message = "Prueba unitaria exitosa: CerrarDocumento Success"
    
TestExit:
    Set Test_CerrarDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function