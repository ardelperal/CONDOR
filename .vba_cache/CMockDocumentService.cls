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

' Variables privadas para configurar el comportamiento del mock
Private m_GenerarDocumento_ReturnValue As String
Private m_LeerDocumento_ReturnValue As Boolean

' Propiedades para verificar si los métodos fueron llamados
Public GenerarDocumento_WasCalled As Boolean
Public LeerDocumento_WasCalled As Boolean

' Propiedades para capturar los últimos parámetros pasados
Private GenerarDocumento_LastSolicitud As ESolicitud
Private LeerDocumento_LastRutaDocumento As String
Private LeerDocumento_LastIdSolicitud As Long

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Implementación de la interfaz IDocumentService

Private Function IDocumentService_GenerarDocumento(ByVal solicitud As ESolicitud) As String
    GenerarDocumento_WasCalled = True
    Set GenerarDocumento_LastSolicitud = solicitud
    IDocumentService_GenerarDocumento = m_GenerarDocumento_ReturnValue
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento_WasCalled = True
    LeerDocumento_LastRutaDocumento = rutaDocumento
    LeerDocumento_LastIdSolicitud = idSolicitud
    IDocumentService_LeerDocumento = m_LeerDocumento_ReturnValue
End Function

' Métodos públicos de conveniencia
Public Function GenerarDocumento(ByVal solicitud As ESolicitud) As String
    GenerarDocumento = IDocumentService_GenerarDocumento(solicitud)
End Function

Public Function LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
    LeerDocumento = IDocumentService_LeerDocumento(rutaDocumento, idSolicitud)
End Function

' Métodos públicos para configurar el mock
Public Sub ConfigureGenerarDocumento(ByVal returnValue As String)
    m_GenerarDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureLeerDocumento(ByVal returnValue As Boolean)
    m_LeerDocumento_ReturnValue = returnValue
End Sub

' Método para resetear el estado del mock
Public Sub Reset()
    m_GenerarDocumento_ReturnValue = ""
    m_LeerDocumento_ReturnValue = False
    GenerarDocumento_WasCalled = False
    LeerDocumento_WasCalled = False
    Set GenerarDocumento_LastSolicitud = Nothing
    LeerDocumento_LastRutaDocumento = ""
    LeerDocumento_LastIdSolicitud = 0
End Sub