Attribute VB_Name = "IntegrationTest_WordManager"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: IntegrationTest_WordManager
' Descripción: Pruebas de integración para CWordManager
' Autor: CONDOR-Expert
' Fecha: 2025-01-15
' Versión: 2.0
' ============================================================================

' Variables para manejo de archivos temporales
Private m_TempFolder As String
Private m_TempFiles As Collection

' ============================================================================
' CONFIGURACIÓN DE PRUEBAS
' ============================================================================

Public Function IntegrationTest_WordManager_RunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "IntegrationTest_WordManager"
    
    On Error GoTo ErrorHandler
    
    ' Inicializar setup centralizado
    Call InicializarSetup
    
    ' Ejecutar pruebas individuales
    suiteResult.AddTestResult IntegrationTest_WordManager_CicloCompleto_Success()
    suiteResult.AddTestResult IntegrationTest_WordManager_AbrirFicheroInexistente_DevuelveFalse()
    
    ' Limpiar recursos
    Call LimpiarArchivosTemporales
    
    Set IntegrationTest_WordManager_RunAll = suiteResult
    Exit Function
    
ErrorHandler:
    Dim errorResult As New CTestResult
    errorResult.Initialize "IntegrationTest_WordManager_RunAll_Setup"
    errorResult.Fail "Error en setup/teardown: " & Err.Number & " - " & Err.Description
    suiteResult.AddTestResult errorResult
    
    Call LimpiarArchivosTemporales
    Set IntegrationTest_WordManager_RunAll = suiteResult
End Function

' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function IntegrationTest_WordManager_CicloCompleto_Success() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "IntegrationTest_WordManager_CicloCompleto_Success"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim archivoOriginal As String
    Dim archivoGuardado As String
    Dim contenidoFinal As String
    
    ' Crear instancia real de ErrorHandler y WordManager usando factory
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    Set config = New CConfig
    Set fileSystem = New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    Set wordManager = modWordManagerFactory.CreateWordManager(errorHandler)
    
    ' Crear archivo de prueba con marcador
    archivoOriginal = m_TempFolder & "documento_original.docx"
    Call CrearDocumentoPrueba(archivoOriginal, "Hola [NOMBRE], este es un documento de prueba.")
    m_TempFiles.Add archivoOriginal
    
    archivoGuardado = m_TempFolder & "documento_modificado.docx"
    m_TempFiles.Add archivoGuardado
    
    ' Act & Assert
    ' 1. Abrir documento
    modAssert.AssertTrue wordManager.AbrirDocumento(archivoOriginal), "Debería abrir el documento correctamente"
    
    ' 2. Reemplazar texto
    modAssert.AssertTrue wordManager.ReemplazarTexto("[NOMBRE]", "CONDOR"), "Debería reemplazar el texto correctamente"
    
    ' 3. Guardar documento
    modAssert.AssertTrue wordManager.GuardarDocumento(archivoGuardado), "Debería guardar el documento correctamente"
    
    ' 4. Cerrar documento
    wordManager.CerrarDocumento
    
    ' 5. Verificar contenido del archivo guardado
    contenidoFinal = wordManager.LeerContenidoDocumento(archivoGuardado)
    modAssert.AssertTrue InStr(contenidoFinal, "CONDOR") > 0, "El contenido debería incluir 'CONDOR'"
    modAssert.AssertTrue InStr(contenidoFinal, "[NOMBRE]") = 0, "El contenido no debería incluir '[NOMBRE]'"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    
    ' Limpiar recursos en caso de error
    On Error Resume Next
    wordManager.CerrarDocumento
    On Error GoTo 0
    
Cleanup:
    ' Limpiar recursos
    On Error Resume Next
    Set wordManager = Nothing
    On Error GoTo 0
    
    Set IntegrationTest_WordManager_CicloCompleto_Success = testResult
End Function

Private Function IntegrationTest_WordManager_AbrirFicheroInexistente_DevuelveFalse() As CTestResult
    Dim testResult As New CTestResult
    testResult.Initialize "IntegrationTest_WordManager_AbrirFicheroInexistente_DevuelveFalse"
    
    On Error GoTo ErrorHandler
    
    ' Arrange
    Dim wordManager As IWordManager
    Dim errorHandler As IErrorHandlerService
    Dim rutaInvalida As String
    Dim resultado As Boolean
    
    ' Crear instancia real de ErrorHandler y WordManager usando factory
    Dim config As IConfig
    Dim fileSystem As IFileSystem
    Set config = New CConfig
    Set fileSystem = New CFileSystem
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService(config, fileSystem)
    Set wordManager = modWordManagerFactory.CreateWordManager(errorHandler)
    
    rutaInvalida = m_TempFolder & "archivo_que_no_existe.docx"
    
    ' Act
    resultado = wordManager.AbrirDocumento(rutaInvalida)
    
    ' Assert
    modAssert.AssertFalse resultado, "Debería devolver False al intentar abrir un archivo inexistente"
    
    testResult.Pass
    GoTo Cleanup
    
ErrorHandler:
    testResult.Fail "Error inesperado: " & Err.Number & " - " & Err.Description
    
Cleanup:
    ' Limpiar recursos
    On Error Resume Next
    Set wordManager = Nothing
    On Error GoTo 0
    
    Set IntegrationTest_WordManager_AbrirFicheroInexistente_DevuelveFalse = testResult
End Function

' ============================================================================
' MÉTODOS DE SETUP Y TEARDOWN CENTRALIZADOS
' ============================================================================

Private Sub InicializarSetup()
    On Error GoTo ErrorHandler
    
    ' Inicializar colección de archivos temporales
    Set m_TempFiles = New Collection
    
    ' Configurar carpeta temporal usando ruta estandarizada del proyecto
    m_TempFolder = modTestUtils.GetProjectPath() & "back\test_env\word_tests\"
    
    ' Crear carpeta temporal si no existe usando IFileSystem
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    If Not fs.FolderExists(m_TempFolder) Then
        fs.CreateFolder m_TempFolder
    End If
    
    Set fs = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "InicializarSetup", "Error en inicialización: " & Err.Description
End Sub

' ============================================================================
' MÉTODOS AUXILIARES
' ============================================================================

Private Sub CrearDocumentoPrueba(ByVal rutaArchivo As String, ByVal contenido As String)
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim wordDoc As Object
    
    ' Crear aplicación Word temporal
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    
    ' Crear nuevo documento
    Set wordDoc = wordApp.Documents.Add
    
    ' Insertar contenido
    wordDoc.content.Text = contenido
    
    ' Guardar como .docx
    wordDoc.SaveAs2 rutaArchivo, 16 ' wdFormatXMLDocument
    
    ' Cerrar y limpiar
    wordDoc.Close False
    wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close False
        Set wordDoc = Nothing
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit
        Set wordApp = Nothing
    End If
    On Error GoTo 0
    
    Err.Raise Err.Number, "CrearDocumentoPrueba", "Error creando documento: " & Err.Description
End Sub

Private Sub LimpiarArchivosTemporales()
    On Error Resume Next
    
    Dim i As Integer
    Dim archivo As String
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Eliminar archivos temporales creados durante las pruebas usando IFileSystem
    If Not m_TempFiles Is Nothing Then
        For i = 1 To m_TempFiles.Count
            archivo = m_TempFiles(i)
            If fs.FileExists(archivo) Then
                fs.DeleteFile archivo
            End If
        Next i
    End If
    
    ' Eliminar carpeta temporal usando exclusivamente IFileSystem service
    If fs.FolderExists(m_TempFolder) Then
        ' Intentar eliminar la carpeta - IFileSystem manejará si está vacía o no
        fs.DeleteFolder m_TempFolder
    End If
    
    Set fs = Nothing
    
    On Error GoTo 0
End Sub