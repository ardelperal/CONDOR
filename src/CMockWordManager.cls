Attribute VB_Name = "CMockWordManager"
Option Compare Database
Option Explicit
Implements IWordManager

' --- Variables de estado para configurar el comportamiento ---
Private m_AbrirDocumento_ReturnValue As Boolean
Private m_ReemplazarTexto_ReturnValue As Boolean
Private m_GuardarDocumento_ReturnValue As Boolean
Private m_LeerDocumento_ReturnValue As String

' --- Propiedades de seguimiento para las aserciones ---
Public AbrirDocumento_WasCalled As Boolean
Public GuardarDocumento_WasCalled As Boolean
Public ReemplazarTexto_CallCount As Long
Public LeerDocumento_CallCount As Long
Public LeerDocumento_FilePath_Param As String

' --- Métodos públicos para configurar el mock ---
Public Sub ConfigureAbrirDocumento(ByVal returnValue As Boolean)
    m_AbrirDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureReemplazarTexto(ByVal returnValue As Boolean)
    m_ReemplazarTexto_ReturnValue = returnValue
End Sub

Public Sub ConfigureGuardarDocumento(ByVal returnValue As Boolean)
    m_GuardarDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureLeerDocumento(ByVal returnValue As String)
    m_LeerDocumento_ReturnValue = returnValue
End Sub

Public Sub Reset()
    m_AbrirDocumento_ReturnValue = False
    m_ReemplazarTexto_ReturnValue = False
    m_GuardarDocumento_ReturnValue = False
    m_LeerDocumento_ReturnValue = ""
    AbrirDocumento_WasCalled = False
    GuardarDocumento_WasCalled = False
    ReemplazarTexto_CallCount = 0
    LeerDocumento_CallCount = 0
    LeerDocumento_FilePath_Param = ""
End Sub

' --- Implementación de la Interfaz ---
Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    AbrirDocumento_WasCalled = True
    IWordManager_AbrirDocumento = m_AbrirDocumento_ReturnValue
End Function

Private Sub IWordManager_CerrarDocumento()
    ' No necesita implementación para el mock
End Sub

Private Function IWordManager_GuardarDocumento(ByVal rutaDestino As String) As Boolean
    GuardarDocumento_WasCalled = True
    IWordManager_GuardarDocumento = m_GuardarDocumento_ReturnValue
End Function

Private Function IWordManager_LeerContenidoDocumento(ByVal rutaDocumento As String) As String
    ' Devuelve un valor por defecto, puede mejorarse si es necesario
    IWordManager_LeerContenidoDocumento = "Contenido de prueba del mock."
End Function

Private Function IWordManager_ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
    ReemplazarTexto_CallCount = ReemplazarTexto_CallCount + 1
    IWordManager_ReemplazarTexto = m_ReemplazarTexto_ReturnValue
End Function

Private Function IWordManager_LeerDocumento(ByVal filePath As String) As String
    LeerDocumento_CallCount = LeerDocumento_CallCount + 1
    LeerDocumento_FilePath_Param = filePath
    IWordManager_LeerDocumento = m_LeerDocumento_ReturnValue
End Function