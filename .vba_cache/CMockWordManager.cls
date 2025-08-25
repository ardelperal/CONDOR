Option Compare Database
Option Explicit
' ============================================================================
' Clase: CMockWordManager
' DescripciÃ³n: Mock del gestor de Word para pruebas unitarias aisladas.
' Implementa: IWordManager
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 1.0
' ============================================================================

Implements IWordManager

' ============================================================================
' PROPIEDADES PÃšBLICAS PARA CONTROL DE COMPORTAMIENTO EN PRUEBAS
' ============================================================================

' Propiedades para AbrirDocumento
Public AbrirDocumento_ReturnValue As Boolean
Public AbrirDocumento_WasCalled As Boolean
Public AbrirDocumento_CallCount As Long
Public AbrirDocumento_LastRutaDocumento As String

' Propiedades para ReemplazarTexto
Public ReemplazarTexto_ReturnValue As Boolean
Public ReemplazarTexto_WasCalled As Boolean
Public ReemplazarTexto_CallCount As Long
Public ReemplazarTexto_LastTextoABuscar As String
Public ReemplazarTexto_LastTextoReemplazo As String
Public ReemplazarTexto_Calls As Collection

' Propiedades para GuardarDocumento
Public GuardarDocumento_ReturnValue As Boolean
Public GuardarDocumento_WasCalled As Boolean
Public GuardarDocumento_CallCount As Long
Public GuardarDocumento_LastRutaDestino As String

' Propiedades para CerrarDocumento
Public CerrarDocumento_WasCalled As Boolean
Public CerrarDocumento_CallCount As Long

' Propiedades para LeerContenidoDocumento
Public LeerContenidoDocumento_ReturnValue As String
Public LeerContenidoDocumento_WasCalled As Boolean
Public LeerContenidoDocumento_CallCount As Long
Public LeerContenidoDocumento_LastRutaDocumento As String

' ============================================================================
' CONSTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Set ReemplazarTexto_Calls = New Collection
    Call Reset
End Sub

' ============================================================================
' MÃ‰TODOS PÃšBLICOS PARA GESTIÃ“N DEL MOCK
' ============================================================================

' Reinicia todas las propiedades del mock a sus valores por defecto
Public Sub Reset()
    ' Reset AbrirDocumento
    AbrirDocumento_ReturnValue = True
    AbrirDocumento_WasCalled = False
    AbrirDocumento_CallCount = 0
    AbrirDocumento_LastRutaDocumento = ""
    
    ' Reset ReemplazarTexto
    ReemplazarTexto_ReturnValue = True
    ReemplazarTexto_WasCalled = False
    ReemplazarTexto_CallCount = 0
    ReemplazarTexto_LastTextoABuscar = ""
    ReemplazarTexto_LastTextoReemplazo = ""
    Set ReemplazarTexto_Calls = New Collection
    
    ' Reset GuardarDocumento
    GuardarDocumento_ReturnValue = True
    GuardarDocumento_WasCalled = False
    GuardarDocumento_CallCount = 0
    GuardarDocumento_LastRutaDestino = ""
    
    ' Reset CerrarDocumento
    CerrarDocumento_WasCalled = False
    CerrarDocumento_CallCount = 0
    
    ' Reset LeerContenidoDocumento
    LeerContenidoDocumento_ReturnValue = ""
    LeerContenidoDocumento_WasCalled = False
    LeerContenidoDocumento_CallCount = 0
    LeerContenidoDocumento_LastRutaDocumento = ""
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE LA INTERFAZ IWordManager
' ============================================================================

Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    AbrirDocumento_WasCalled = True
    AbrirDocumento_CallCount = AbrirDocumento_CallCount + 1
    AbrirDocumento_LastRutaDocumento = rutaDocumento
    IWordManager_AbrirDocumento = AbrirDocumento_ReturnValue
End Function

Private Function IWordManager_ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
    ReemplazarTexto_WasCalled = True
    ReemplazarTexto_CallCount = ReemplazarTexto_CallCount + 1
    ReemplazarTexto_LastTextoABuscar = textoABuscar
    ReemplazarTexto_LastTextoReemplazo = textoReemplazo
    
    ' Guardar la llamada en la colecciÃ³n para verificaciÃ³n detallada
    Dim callInfo As String
    callInfo = textoABuscar & "|" & textoReemplazo
    ReemplazarTexto_Calls.Add callInfo
    
    IWordManager_ReemplazarTexto = ReemplazarTexto_ReturnValue
End Function

Private Function IWordManager_GuardarDocumento(ByVal rutaDestino As String) As Boolean
    GuardarDocumento_WasCalled = True
    GuardarDocumento_CallCount = GuardarDocumento_CallCount + 1
    GuardarDocumento_LastRutaDestino = rutaDestino
    IWordManager_GuardarDocumento = GuardarDocumento_ReturnValue
End Function

Private Sub IWordManager_CerrarDocumento()
    CerrarDocumento_WasCalled = True
    CerrarDocumento_CallCount = CerrarDocumento_CallCount + 1
End Sub

Private Function IWordManager_LeerContenidoDocumento(ByVal rutaDocumento As String) As String
    LeerContenidoDocumento_WasCalled = True
    LeerContenidoDocumento_CallCount = LeerContenidoDocumento_CallCount + 1
    LeerContenidoDocumento_LastRutaDocumento = rutaDocumento
    IWordManager_LeerContenidoDocumento = LeerContenidoDocumento_ReturnValue
End Function
