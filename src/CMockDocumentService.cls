VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockDocumentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' Clase: CMockDocumentService
' Descripción: Mock del servicio de documentos para pruebas unitarias aisladas.
' Implementa: IDocumentService
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' Versión: 1.0
' ============================================================================

Implements IDocumentService

' ============================================================================
' PROPIEDADES PÚBLICAS PARA CONTROL DE COMPORTAMIENTO EN PRUEBAS
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
' MÉTODOS PÚBLICOS PARA GESTIÓN DEL MOCK
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
' IMPLEMENTACIÓN DE LA INTERFAZ IDocumentService
' ============================================================================

Private Function IDocumentService_GenerarDocumento(ByVal solicitud As T_Solicitud) As String
    GenerarDocumento_WasCalled = True
    GenerarDocumento_CallCount = GenerarDocumento_CallCount + 1
    Set GenerarDocumento_LastSolicitud = solicitud
    IDocumentService_GenerarDocumento = GenerarDocumento_ReturnValue
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento_WasCalled = True
    LeerDocumento_CallCount = LeerDocumento_CallCount + 1
    LeerDocumento_LastRutaDocumento = rutaDocumento
    LeerDocumento_LastIdSolicitud = idSolicitud
    IDocumentService_LeerDocumento = LeerDocumento_ReturnValue
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para GenerarDocumento
Public Function GenerarDocumento(ByVal solicitud As T_Solicitud) As String
    GenerarDocumento = IDocumentService_GenerarDocumento(solicitud)
End Function

' Método público de conveniencia para LeerDocumento
Public Function LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento = IDocumentService_LeerDocumento(rutaDocumento, idSolicitud)
End Function


