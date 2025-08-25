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
    suite.AddTest Test_AbrirYGuardarDocumento_ConRutaValida_DebeEjecutarCorrectamente
    suite.AddTest Test_AbrirDocumento_ConRutaInexistente_DebeRetornarFalse
    suite.AddTest Test_LeerContenidoDocumento_ConDocumentoValido_DebeRetornarContenido
    suite.AddTest Test_CerrarDocumento_ConDocumentoAbierto_DebeEjecutarSinErrores
    
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
    
    Dim wordManager As New CWordManager
    Dim rutaPlantilla As String
    Dim rutaDestino As String
    Dim fso As Object
    
    ' Configurar rutas de prueba
    rutaPlantilla = "C:\Temp\plantilla_test.docx"
    rutaDestino = "C:\Temp\documento_generado_test.docx"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear archivo de plantilla de prueba si no existe
    If Not fso.FileExists(rutaPlantilla) Then
        CrearArchivoPlantillaPrueba rutaPlantilla
    End If
    
    ' Limpiar archivo destino si existe
    If fso.FileExists(rutaDestino) Then
        fso.DeleteFile rutaDestino
    End If
    
    ' Ejecutar operaciones
    Dim abrirResult As Boolean
    Dim reemplazarResult As Boolean
    Dim guardarResult As Boolean
    
    abrirResult = wordManager.AbrirDocumento(rutaPlantilla)
    modAssert.AssertTrue abrirResult, "Debería abrir el documento correctamente"
    
    reemplazarResult = wordManager.ReemplazarTexto("[MARCADOR_PRUEBA]", "TEXTO_REEMPLAZADO")
    modAssert.AssertTrue reemplazarResult, "Debería reemplazar el texto correctamente"
    
    guardarResult = wordManager.GuardarDocumento(rutaDestino)
    modAssert.AssertTrue guardarResult, "Debería guardar el documento correctamente"
    
    ' Verificar que el archivo se creó
    modAssert.AssertTrue fso.FileExists(rutaDestino), "El archivo guardado debería existir"
    
    ' Limpiar recursos
    wordManager.CerrarDocumento
    
    ' Limpiar archivos de prueba
    LimpiarArchivosTemporales
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Abrir, reemplazar, guardar y verificar archivo"
    
TestExit:
    Set Test_AbrirReemplazarGuardar_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    LimpiarArchivosTemporales
    Resume TestExit
End Function

' Prueba que intenta abrir un archivo inexistente y verifica que falla
Private Function Test_AbrirDocumento_Inexistente_Fail() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_AbrirDocumento_Inexistente_Fail"
    
    On Error GoTo TestError
    
    Dim wordManager As New CWordManager
    Dim rutaInexistente As String
    Dim resultado As Boolean
    
    rutaInexistente = "C:\Temp\archivo_que_no_existe_" & Format(Now, "yyyymmddhhnnss") & ".docx"
    
    ' Intentar abrir archivo inexistente
    resultado = wordManager.AbrirDocumento(rutaInexistente)
    
    ' Verificar que devuelve False
    modAssert.AssertFalse resultado, "Debería devolver False al intentar abrir archivo inexistente"
    
    ' Limpiar recursos
    wordManager.CerrarDocumento
    
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
    
    Dim wordManager As New CWordManager
    Dim rutaDocumento As String
    Dim contenido As String
    
    rutaDocumento = "C:\Temp\documento_lectura_test.docx"
    
    ' Crear documento de prueba
    CrearDocumentoPruebaConTexto rutaDocumento, "Contenido de prueba para lectura"
    
    ' Leer contenido
    contenido = wordManager.LeerContenidoDocumento(rutaDocumento)
    
    ' Verificar que se leyó contenido
    modAssert.AssertTrue Len(contenido) > 0, "Debería leer contenido del documento"
    modAssert.AssertTrue InStr(contenido, "Contenido de prueba") > 0, "Debería contener el texto esperado"
    
    ' Limpiar archivos de prueba
    LimpiarArchivosTemporales
    
    testResult.Passed = True
    testResult.Message = "Prueba exitosa: Lectura de contenido de documento"
    
TestExit:
    Set Test_LeerContenidoDocumento_Success = testResult
    Exit Function
    
TestError:
    testResult.Passed = False
    testResult.Message = "Error en prueba: " & Err.Description
    LimpiarArchivosTemporales
    Resume TestExit
End Function

' Prueba el cierre correcto de documentos
Private Function Test_CerrarDocumento_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.TestName = "Test_CerrarDocumento_Success"
    
    On Error GoTo TestError
    
    Dim wordManager As New CWordManager
    Dim rutaDocumento As String
    
    rutaDocumento = "C:\Temp\documento_cierre_test.docx"
    
    ' Crear y abrir documento
    CrearArchivoPlantillaPrueba rutaDocumento
    wordManager.AbrirDocumento rutaDocumento
    
    ' Cerrar documento (no debería generar errores)
    wordManager.CerrarDocumento
    
    ' Limpiar archivos de prueba
    LimpiarArchivosTemporales
    
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