Option Compare Database
Option Explicit
' ============================================================================
' Interfaz: IDocumentService
' DescripciÃ³n: Define el contrato para el servicio de lectura y escritura de documentos.
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 2.1 - Eliminado Initialize de la interfaz.
' ============================================================================

' ============================================================================
' MÃ‰TODOS PRINCIPALES DEL SERVICIO
' ============================================================================

' Genera un documento oficial basado en una solicitud
' ParÃ¡metros:
'   - laSolicitud: El objeto de solicitud que contiene todos los datos.
' Retorna: Ruta completa del documento generado o una cadena vacÃ­a si falla.
Public Function GenerarDocumento(ByVal laSolicitud As T_Solicitud) As String
End Function

' Lee un documento Word y actualiza la solicitud correspondiente en la base de datos.
' ParÃ¡metros:
'   - rutaDocumento: La ruta completa al fichero Word que se va a leer.
'   - idSolicitud: El ID de la solicitud que se debe actualizar con los datos del documento.
' Retorna: True si la lectura y actualizaciÃ³n fueron exitosas, False en caso contrario.
Public Function LeerDocumento(ByVal rutaDocumento As String, ByVal idSolicitud As Long) As Boolean
End Function
