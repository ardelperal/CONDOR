Option Compare Database
Option Explicit
' ============================================================================
' MÃ³dulo: Test_WordManager
' DescripciÃ³n: Suite de pruebas para WordManager con pruebas de integraciÃ³n controladas.
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 1.0
' ============================================================================

' ============================================================================
' FUNCIÃ“N PRINCIPAL DE EJECUCIÃ“N
' ============================================================================

Public Function Test_WordManager_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.SuiteName = "Test_WordManager"
    
    ' Ejecutar todas las pruebas
    suiteResult.AddTestResult Test_AbrirReemplazarGuardar_Success()
    suiteResult.AddTestResult Test_AbrirDocumento_Inexistente_Fail()
    suiteResult.AddTestResult Test_LeerContenidoDocumento_Success()
    suiteResult.AddTestResult Test_CerrarDocumento_Success()
    
    Set Test_WordManager_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÃ“N CONTROLADAS
' ============================================================================

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
    modAssert.AssertTrue abrirResult, "DeberÃ­a abrir el documento correctamente"
    
    reemplazarResult = wordManager.ReemplazarTexto("[MARCADOR_PRUEBA]", "TEXTO_REEMPLAZADO")
    modAssert.AssertTrue reemplazarResult, "DeberÃ­a reemplazar el texto correctamente"
    
    guardarResult = wordManager.GuardarDocumento(rutaDestino)
    modAssert.AssertTrue guardarResult, "DeberÃ­a guardar el documento correctamente"
    
    ' Verificar que el archivo se creÃ³
    modAssert.AssertTrue fso.FileExists(rutaDestino), "El archivo guardado deberÃ­a existir"
    
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
    modAssert.AssertFalse resultado, "DeberÃ­a devolver False al intentar abrir archivo inexistente"
    
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
    
    ' Verificar que se leyÃ³ contenido
    modAssert.AssertTrue Len(contenido) > 0, "DeberÃ­a leer contenido del documento"
    modAssert.AssertTrue InStr(contenido, "Contenido de prueba") > 0, "DeberÃ­a contener el texto esperado"
    
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
    
    ' Cerrar documento (no deberÃ­a generar errores)
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
' MÃ‰TODOS AUXILIARES PARA PRUEBAS
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

' Crea un documento con texto especÃ­fico para pruebas de lectura
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
