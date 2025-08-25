Option Compare Database
Option Explicit
' ============================================================================
' Clase: CWordManager
' Descripción: Implementación concreta del gestor de Word.
' Implementa: IWordManager
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' Versión: 1.0
' ============================================================================

Implements IWordManager

' Variables privadas para manejo de Word
Private m_WordApp As Object
Private m_WordDoc As Object
Private m_IsWordAppCreated As Boolean

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    m_IsWordAppCreated = False
    Set m_WordApp = Nothing
    Set m_WordDoc = Nothing
End Sub

Private Sub Class_Terminate()
    Call LimpiarRecursos
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IWordManager
' ============================================================================

Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    
    ' Crear aplicación Word si no existe
    If m_WordApp Is Nothing Then
        Set m_WordApp = CreateObject("Word.Application")
        m_IsWordAppCreated = True
        m_WordApp.Visible = False
    End If
    
    ' Abrir el documento
    Set m_WordDoc = m_WordApp.Documents.Open(rutaDocumento)
    
    IWordManager_AbrirDocumento = True
    Exit Function
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CWordManager.AbrirDocumento"
    IWordManager_AbrirDocumento = False
End Function

Private Function IWordManager_ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        errorHandler.LogError -1, "No hay documento abierto", "CWordManager.ReemplazarTexto"
        IWordManager_ReemplazarTexto = False
        Exit Function
    End If
    
    ' Realizar el reemplazo
    With m_WordDoc.Content.Find
        .Text = textoABuscar
        .Replacement.Text = textoReemplazo
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    IWordManager_ReemplazarTexto = True
    Exit Function
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CWordManager.ReemplazarTexto"
    IWordManager_ReemplazarTexto = False
End Function

Private Function IWordManager_GuardarDocumento(ByVal rutaDestino As String) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    If m_WordDoc Is Nothing Then
        Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
        errorHandler.LogError -1, "No hay documento abierto", "CWordManager.GuardarDocumento"
        IWordManager_GuardarDocumento = False
        Exit Function
    End If
    
    ' Guardar el documento
    m_WordDoc.SaveAs2 rutaDestino
    
    IWordManager_GuardarDocumento = True
    Exit Function
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CWordManager.GuardarDocumento"
    IWordManager_GuardarDocumento = False
End Function

Private Sub IWordManager_CerrarDocumento()
    Call LimpiarRecursos
End Sub

Private Function IWordManager_LeerContenidoDocumento(ByVal rutaDocumento As String) As String
    Dim errorHandler As IErrorHandlerService
    Dim tempWordApp As Object
    Dim tempWordDoc As Object
    Dim contenido As String
    On Error GoTo ErrorHandler
    
    ' Crear aplicación temporal para lectura
    Set tempWordApp = CreateObject("Word.Application")
    tempWordApp.Visible = False
    
    ' Abrir documento
    Set tempWordDoc = tempWordApp.Documents.Open(rutaDocumento)
    
    ' Leer contenido
    contenido = tempWordDoc.Content.Text
    
    ' Limpiar recursos temporales
    tempWordDoc.Close False
    tempWordApp.Quit
    Set tempWordDoc = Nothing
    Set tempWordApp = Nothing
    
    IWordManager_LeerContenidoDocumento = contenido
    Exit Function
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CWordManager.LeerContenidoDocumento"
    
    ' Limpiar recursos en caso de error
    If Not tempWordDoc Is Nothing Then
        tempWordDoc.Close False
        Set tempWordDoc = Nothing
    End If
    If Not tempWordApp Is Nothing Then
        tempWordApp.Quit
        Set tempWordApp = Nothing
    End If
    
    IWordManager_LeerContenidoDocumento = ""
End Function

' ============================================================================
' MÉTODOS PRIVADOS
' ============================================================================

Private Sub LimpiarRecursos()
    On Error Resume Next
    
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
    End If
    
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        m_WordApp.Quit
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
    End If
End Sub