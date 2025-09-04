VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

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
Private Sub IWordManager_Dispose()
    ' En un mock, este método puede estar vacío o registrar que fue llamado.
    ' Por ahora, lo dejamos vacío para resolver la compilación.
End Sub
