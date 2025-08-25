Option Compare Database
Option Explicit
' ============================================================================
' Clase: CMockDocumentService
' DescripciÃ³n: Mock del servicio de documentos para pruebas unitarias aisladas.
' Implementa: IDocumentService
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 1.0
' ============================================================================

Implements IDocumentService

' ============================================================================
' PROPIEDADES PÃšBLICAS PARA CONTROL DE COMPORTAMIENTO EN PRUEBAS
' ============================================================================

' Propiedades para GenerarDocumento
Public GenerarDocumento_ReturnValue As String
Public GenerarDocumento_WasCalled As Boolean
Public GenerarDocumento_CallCount As Long
Public GenerarDocumento_LastSolicitud As T_Solicitud

' Propiedades para LeerDocumento
Public LeerDocumento_ReturnValue As Boolean
Public LeerDocumento_WasCalled As Boolean
Public LeerDocumento_CallCount As Long
Public LeerDocumento_LastRutaDocumento As String
Public LeerDocumento_LastIdSolicitud As Long

' ============================================================================
' MÃ‰TODOS PÃšBLICOS PARA GESTIÃ“N DEL MOCK
' ============================================================================

' Reinicia todas las propiedades del mock a sus valores por defecto
Public Sub Reset()
    ' Reset GenerarDocumento
    GenerarDocumento_ReturnValue = ""
    GenerarDocumento_WasCalled = False
    GenerarDocumento_CallCount = 0
    Set GenerarDocumento_LastSolicitud = Nothing
    
    ' Reset LeerDocumento
    LeerDocumento_ReturnValue = False
    LeerDocumento_WasCalled = False
    LeerDocumento_CallCount = 0
    LeerDocumento_LastRutaDocumento = ""
    LeerDocumento_LastIdSolicitud = 0
End Sub

' ============================================================================
' IMPLEMENTACIÃ“N DE LA INTERFAZ IDocumentService
' ============================================================================

Private Function IDocumentService_GenerarDocumento(ByVal laSolicitud As T_Solicitud) As String
    GenerarDocumento_WasCalled = True
    GenerarDocumento_CallCount = GenerarDocumento_CallCount + 1
    Set GenerarDocumento_LastSolicitud = laSolicitud
    IDocumentService_GenerarDocumento = GenerarDocumento_ReturnValue
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento_WasCalled = True
    LeerDocumento_CallCount = LeerDocumento_CallCount + 1
    LeerDocumento_LastRutaDocumento = rutaDocumento
    LeerDocumento_LastIdSolicitud = idSolicitud
    IDocumentService_LeerDocumento = LeerDocumento_ReturnValue
End Function
