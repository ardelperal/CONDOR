VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private m_ErrorHandler As IErrorHandlerService ' Inyectar ErrorHandler

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
' INICIALIZACIÓN CON DEPENDENCIAS
' ============================================================================
Public Sub Initialize(ByVal wordApp As Object, ByVal errorHandler As IErrorHandlerService)
    Set m_WordApp = wordApp
    Set m_ErrorHandler = errorHandler
    m_IsWordAppCreated = False ' La aplicación Word se inyecta, no se crea internamente
End Sub

' ============================================================================
' IMPLEMENTACIÓN DE LA INTERFAZ IWordManager
' ============================================================================

Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    On Error GoTo errorHandler
    
    ' Verificar que la aplicación Word esté disponible
    If m_WordApp Is Nothing Then
        If Not m_ErrorHandler Is Nothing Then
            m_ErrorHandler.LogError -1, "Aplicación Word no inicializada", "CWordManager.IWordManager_AbrirDocumento"
        End If
        IWordManager_AbrirDocumento = False
        Exit Function
    End If
    
    ' Abrir el documento
    Set m_WordDoc = m_WordApp.Documents.Open(rutaDocumento)
    
    IWordManager_AbrirDocumento = True
    Exit Function
    
errorHandler:
    ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CWordManager.IWordManager_AbrirDocumento"
    End If
    
    ' Limpiar recursos en caso de error
    If Not m_WordDoc Is Nothing Then
        On Error Resume Next
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
        On Error GoTo 0
    End If
    
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        On Error Resume Next
        m_WordApp.Quit
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
        On Error GoTo 0
    End If
    
    IWordManager_AbrirDocumento = False
End Function

Private Function IWordManager_ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
    On Error GoTo errorHandler
    
    If m_WordDoc Is Nothing Then
        ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
        If Not m_ErrorHandler Is Nothing Then
            m_ErrorHandler.LogError -1, "No hay documento abierto", "CWordManager.IWordManager_ReemplazarTexto"
        End If
        IWordManager_ReemplazarTexto = False
        Exit Function
    End If
    
    ' Realizar el reemplazo
    With m_WordDoc.content.Find
        .Text = textoABuscar
        .Replacement.Text = textoReemplazo
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    IWordManager_ReemplazarTexto = True
    Exit Function
    
errorHandler:
    ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CWordManager.IWordManager_ReemplazarTexto"
    End If
    IWordManager_ReemplazarTexto = False
End Function

Private Function IWordManager_GuardarDocumento(ByVal rutaDestino As String) As Boolean
    On Error GoTo errorHandler
    
    If m_WordDoc Is Nothing Then
        ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
        If Not m_ErrorHandler Is Nothing Then
            m_ErrorHandler.LogError -1, "No hay documento abierto", "CWordManager.IWordManager_GuardarDocumento"
        End If
        IWordManager_GuardarDocumento = False
        Exit Function
    End If
    
    ' Guardar el documento
    m_WordDoc.SaveAs2 rutaDestino
    
    IWordManager_GuardarDocumento = True
    Exit Function
    
errorHandler:
    ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CWordManager.IWordManager_GuardarDocumento"
    End If
    IWordManager_GuardarDocumento = False
End Function

Private Sub IWordManager_CerrarDocumento()
    Call LimpiarRecursos
End Sub

Private Function IWordManager_LeerContenidoDocumento(ByVal rutaDocumento As String) As String
    Dim contenido As String
    Dim tempWordDoc As Object
    
    On Error GoTo errorHandler
    
    ' Verificar que la aplicación Word esté disponible
    If m_WordApp Is Nothing Then
        If Not m_ErrorHandler Is Nothing Then
            m_ErrorHandler.LogError -1, "Aplicación Word no inicializada", "CWordManager.IWordManager_LeerContenidoDocumento"
        End If
        IWordManager_LeerContenidoDocumento = ""
        Exit Function
    End If
    
    ' Abrir el documento usando la aplicación Word inyectada
    Set tempWordDoc = m_WordApp.Documents.Open(rutaDocumento)
    
    ' Leer el contenido
    contenido = tempWordDoc.content.Text
    
    ' Limpiar recursos temporales
    tempWordDoc.Close False
    Set tempWordDoc = Nothing
    
    IWordManager_LeerContenidoDocumento = contenido
    Exit Function
    
errorHandler:
    ' Corregir llamada a LogError (Lección 1) y usar m_ErrorHandler (Lección 8)
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CWordManager.IWordManager_LeerContenidoDocumento"
    End If
    
    ' Limpiar recursos en caso de error
    If Not tempWordDoc Is Nothing Then
        On Error Resume Next
        tempWordDoc.Close False
        Set tempWordDoc = Nothing
        On Error GoTo 0
    End If
    
    IWordManager_LeerContenidoDocumento = ""
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA (Lección 24)
' ============================================================================

Public Function AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    AbrirDocumento = IWordManager_AbrirDocumento(rutaDocumento)
End Function

Public Function ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
    ReemplazarTexto = IWordManager_ReemplazarTexto(textoABuscar, textoReemplazo)
End Function

Public Function GuardarDocumento(ByVal rutaDestino As String) As Boolean
    GuardarDocumento = IWordManager_GuardarDocumento(rutaDestino)
End Function

Public Sub CerrarDocumento()
    Call IWordManager_CerrarDocumento
End Sub

Public Function LeerContenidoDocumento(ByVal rutaDocumento As String) As String
    LeerContenidoDocumento = IWordManager_LeerContenidoDocumento(rutaDocumento)
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

Private Sub IWordManager_Dispose()
    On Error Resume Next
    ' Cierra docs y la app
    If Not m_WordApp Is Nothing Then
        Do While m_WordApp.Documents.Count > 0
            m_WordApp.Documents(1).Close False
        Loop
        m_WordApp.Quit 0
    End If
    Set m_WordApp = Nothing
    ' Fallback (solo en testing): limpiar posibles zómbis recientes
    Call modTestUtils.KillRecentWordProcesses(10, True)
End Sub

