Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Interfaz: IDocumentService
' Versión: 3.0 - Alinear con la arquitectura de servicios colaborativos
' ============================================================================
Public Function GenerarDocumento(ByVal SolicitudID As Long) As String
End Function

Public Function LeerDocumento(ByVal rutaDocumento As String) As ESolicitud
End Function