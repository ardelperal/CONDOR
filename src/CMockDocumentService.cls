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

Implements IDocumentService

' Propiedades para configurar el comportamiento del mock
Public GenerarDocumento_ReturnValue As String
Public LeerDocumento_ReturnValue As Boolean

' Propiedades para verificar si los métodos fueron llamados
Public GenerarDocumento_WasCalled As Boolean
Public LeerDocumento_WasCalled As Boolean

' Propiedades para capturar los últimos parámetros pasados
Private GenerarDocumento_LastSolicitud As ESolicitud
Private LeerDocumento_LastRutaDocumento As String
Private LeerDocumento_LastIdSolicitud As Long

' Implementación de la interfaz IDocumentService
Private Sub IDocumentService_Initialize(ByVal configService As IConfig, ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger, ByVal wordManager As IWordManager, ByVal mapeoRepository As IMapeoRepository, ByVal errorHandler As IErrorHandlerService)
    ' No se necesita inicialización real para el mock
End Sub

Private Function IDocumentService_GenerarDocumento(ByVal solicitud As ESolicitud) As String
    GenerarDocumento_WasCalled = True
    Set GenerarDocumento_LastSolicitud = solicitud
    IDocumentService_GenerarDocumento = GenerarDocumento_ReturnValue
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento_WasCalled = True
    LeerDocumento_LastRutaDocumento = rutaDocumento
    LeerDocumento_LastIdSolicitud = idSolicitud
    IDocumentService_LeerDocumento = LeerDocumento_ReturnValue
End Function

' Métodos públicos de conveniencia
Public Function GenerarDocumento(ByVal solicitud As ESolicitud) As String
    GenerarDocumento = IDocumentService_GenerarDocumento(solicitud)
End Function

Public Function LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento = IDocumentService_LeerDocumento(rutaDocumento, idSolicitud)
End Function

' Método para resetear el estado del mock
Public Sub Reset()
    GenerarDocumento_ReturnValue = ""
    LeerDocumento_ReturnValue = False
    GenerarDocumento_WasCalled = False
    LeerDocumento_WasCalled = False
    Set GenerarDocumento_LastSolicitud = Nothing
    LeerDocumento_LastRutaDocumento = ""
    LeerDocumento_LastIdSolicitud = 0
End Sub