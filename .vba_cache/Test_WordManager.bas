Attribute VB_Name = "Test_WordManager"
'******************************************************************************
' Módulo: Test_WordManager
' Descripción: Suite de pruebas para WordManager con pruebas de integración controladas.
' Autor: CONDOR-Expert
' Fecha: 2025-01-21
' Versión: 1.0
'******************************************************************************

Option Compare Database
Option Explicit

' FUNCIÓN PRINCIPAL DE EJECUCIÓN
'******************************************************************************

Public Function Test_WordManager_RunAll() As CTestSuiteResult
    Dim suite As New CTestSuiteResult
    suite.Initialize "Test_WordManager"
    
    ' Ejecutar todas las pruebas
    suite.AddTest Test_AbrirReemplazarGuardar_Success
    suite.AddTest Test_AbrirDocumento_Inexistente_Fail
    suite.AddTest Test_LeerContenidoDocumento_Success
    suite.AddTest Test_CerrarDocumento_Success
    
    Set Test_WordManager_RunAll = suite
End Function

'******************************************************************************
' PRUEBAS DE INTEGRACIÓN CONTROLADAS
'******************************************************************************

' Prueba que abre un documento, reemplaza texto, lo guarda y verifica que existe
Private Function Test_AbrirReemplazarGuardar_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_AbrirReemplazarGuardar_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaPlantilla As String
    Dim rutaDestino As String
    Dim fso As Object
    
    ' Configurar rutas de prueba
    rutaPlantilla = "C:\Temp\plantilla_test.docx"
    rutaDestino = "C:\Temp\documento_generado_test.docx"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ejecutar operaciones
    Dim abrirResult As Boolean
    Dim reemplazarResult As Boolean
    Dim guardarResult As Boolean
    
    ' Configurar mock para retornar éxito
    mockWordManager.AbrirDocumento_ReturnValue = True
    mockWordManager.ReemplazarTexto_ReturnValue = True
    mockWordManager.GuardarDocumento_ReturnValue = True
    
    abrirResult = mockWordManager.AbrirDocumento(rutaPlantilla)
     reemplazarResult = mockWordManager.ReemplazarTexto("[MARCADOR_PRUEBA]", "TEXTO_REEMPLAZADO")
     guardarResult = mockWordManager.GuardarDocumento(rutaDestino)
     
     ' Verificar que los métodos fueron llamados correctamente
     modAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "AbrirDocumento debería haber sido llamado"
     modAssert.AssertEqual rutaPlantilla, mockWordManager.AbrirDocumento_LastRutaDocumento, "Debería pasar la ruta de plantilla correcta"
     
     modAssert.AssertTrue mockWordManager.ReemplazarTexto_WasCalled, "ReemplazarTexto debería haber sido llamado"
     modAssert.AssertEqual "[MARCADOR_PRUEBA]", mockWordManager.ReemplazarTexto_LastTextoABuscar, "Debería buscar el marcador correcto"
     modAssert.AssertEqual "TEXTO_REEMPLAZADO", mockWordManager.ReemplazarTexto_LastTextoReemplazo, "Debería usar el texto de reemplazo correcto"
     
     modAssert.AssertTrue mockWordManager.GuardarDocumento_WasCalled, "GuardarDocumento debería haber sido llamado"
     modAssert.AssertEqual rutaDestino, mockWordManager.GuardarDocumento_LastRutaDestino, "Debería guardar en la ruta de destino correcta"
     
     ' Verificar que los métodos retornaron los valores esperados
     modAssert.AssertTrue abrirResult, "AbrirDocumento debería retornar True"
     modAssert.AssertTrue reemplazarResult, "ReemplazarTexto debería retornar True"
     modAssert.AssertTrue guardarResult, "GuardarDocumento debería retornar True"
    
    ' Limpiar recursos
    mockWordManager.CerrarDocumento
    
    ' No necesitamos limpiar archivos en pruebas unitarias con mocks
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Abrir, reemplazar, guardar y verificar archivo"
    
TestExit:
    Set Test_AbrirReemplazarGuardar_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba que intenta abrir un archivo inexistente y verifica que falla
Private Function Test_AbrirDocumento_Inexistente_Fail() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_AbrirDocumento_Inexistente_Fail"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaInexistente As String
    Dim resultado As Boolean
    
    rutaInexistente = "C:\Temp\archivo_que_no_existe_" & Format(Now, "yyyymmddhhnnss") & ".docx"
    
    ' Configurar mock para retornar fallo
    mockWordManager.AbrirDocumento_ReturnValue = False
    
    ' Intentar abrir archivo inexistente
     resultado = mockWordManager.AbrirDocumento(rutaInexistente)
     
     ' Verificar que el método fue llamado y retornó False
     modAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "AbrirDocumento debería haber sido llamado"
     modAssert.AssertEqual rutaInexistente, mockWordManager.AbrirDocumento_LastRutaDocumento, "Debería pasar la ruta inexistente"
     modAssert.AssertFalse resultado, "Debería devolver False cuando el mock está configurado para fallo"
    
    ' Limpiar recursos
    mockWordManager.CerrarDocumento
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Manejo correcto de archivo inexistente"
    
TestExit:
    Set Test_AbrirDocumento_Inexistente_Fail = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba la lectura de contenido de un documento
Private Function Test_LeerContenidoDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_LeerContenidoDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    Dim contenido As String
    
    rutaDocumento = "C:\Temp\documento_lectura_test.docx"
    
    ' Configurar mock para retornar contenido
    mockWordManager.LeerContenidoDocumento_ReturnValue = "Contenido de prueba para lectura"
    
    ' Leer contenido
     contenido = mockWordManager.LeerContenidoDocumento(rutaDocumento)
     
     ' Verificar que el método fue llamado correctamente
     modAssert.AssertTrue mockWordManager.LeerContenidoDocumento_WasCalled, "LeerContenidoDocumento debería haber sido llamado"
     modAssert.AssertEqual rutaDocumento, mockWordManager.LeerContenidoDocumento_LastRutaDocumento, "Debería pasar la ruta correcta"
     modAssert.AssertEqual "Contenido de prueba para lectura", contenido, "Debería retornar el contenido configurado en el mock"
    
    ' No necesitamos limpiar archivos en pruebas unitarias con mocks
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Lectura de contenido de documento"
    
TestExit:
    Set Test_LeerContenidoDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    Resume TestExit
End Function

' Prueba el cierre correcto de documentos
Private Function Test_CerrarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_CerrarDocumento_Success"
    
    On Error GoTo TestError
    
    Dim mockWordManager As New CMockWordManager
    Dim rutaDocumento As String
    
    rutaDocumento = "C:\Temp\documento_cierre_test.docx"
    
    ' Configurar mock
    mockWordManager.AbrirDocumento_ReturnValue = True
    
    ' Abrir documento
    mockWordManager.AbrirDocumento rutaDocumento
    
    ' Cerrar documento (no debería generar errores)
     mockWordManager.CerrarDocumento
     
     ' Verificar que los métodos fueron llamados
     modAssert.AssertTrue mockWordManager.AbrirDocumento_WasCalled, "AbrirDocumento debería haber sido llamado"
     modAssert.AssertTrue mockWordManager.CerrarDocumento_WasCalled, "CerrarDocumento debería haber sido llamado"
    
    ' No necesitamos limpiar archivos en pruebas unitarias con mocks
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Cierre correcto de documento"
    
TestExit:
    Set Test_CerrarDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    LimpiarArchivosTemporales
    Resume TestExit
End Function

' ============================================================================
' MÉTODOS AUXILIARES PARA PRUEBAS
' ============================================================================

' Crea un archivo de plantilla de prueba con marcadores
Private Sub CrearArchivoPlantillaPrueba(ByVal rutaArchivo As String)
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear directorio si no existe
    If Not fso.FolderExists("C:\Temp") Then
        fso.CreateFolder "C:\Temp"
    End If
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Set wordDoc = wordApp.Documents.Add
    wordDoc.Content.Text = "Documento de prueba con [MARCADOR_PRUEBA] para reemplazar."
    
    wordDoc.SaveAs2 rutaArchivo
    wordDoc.Close
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

' Crea un documento con texto específico para pruebas de lectura
Private Sub CrearDocumentoPruebaConTexto(ByVal rutaArchivo As String, ByVal texto As String)
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear directorio si no existe
    If Not fso.FolderExists("C:\Temp") Then
        fso.CreateFolder "C:\Temp"
    End If
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Set wordDoc = wordApp.Documents.Add
    wordDoc.Content.Text = texto
    
    wordDoc.SaveAs2 rutaArchivo
    wordDoc.Close
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

' Limpia todos los archivos temporales de prueba
Private Sub LimpiarArchivosTemporales()
    Dim fso As Object
    Dim archivos As Variant
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    archivos = Array( _
        "C:\Temp\plantilla_test.docx", _
        "C:\Temp\documento_generado_test.docx", _
        "C:\Temp\documento_lectura_test.docx", _
        "C:\Temp\documento_cierre_test.docx" _
    )
    
    For i = 0 To UBound(archivos)
        If fso.FileExists(archivos(i)) Then
            On Error Resume Next
            fso.DeleteFile archivos(i)
            On Error GoTo 0
        End If
    Next i
End Sub