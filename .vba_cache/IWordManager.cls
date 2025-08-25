Option Compare Database
Option Explicit
' ============================================================================
' Interfaz: IWordManager
' DescripciÃ³n: Define el contrato para el manejo de documentos Word.
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' VersiÃ³n: 1.0
' ============================================================================

' ============================================================================
' MÃ‰TODOS PARA MANEJO DE WORD
' ============================================================================

' Abre un documento Word desde una ruta especÃ­fica
' ParÃ¡metros:
'   - rutaDocumento: Ruta completa al archivo Word
' Retorna: True si se abriÃ³ correctamente, False en caso contrario
Public Function AbrirDocumento(ByVal rutaDocumento As String) As Boolean
End Function

' Reemplaza texto en el documento actualmente abierto
' ParÃ¡metros:
'   - textoABuscar: Texto que se va a reemplazar
'   - textoReemplazo: Texto de reemplazo
' Retorna: True si se realizÃ³ el reemplazo, False en caso contrario
Public Function ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
End Function

' Guarda el documento actualmente abierto en una ruta especÃ­fica
' ParÃ¡metros:
'   - rutaDestino: Ruta donde guardar el documento
' Retorna: True si se guardÃ³ correctamente, False en caso contrario
Public Function GuardarDocumento(ByVal rutaDestino As String) As Boolean
End Function

' Cierra el documento y libera los recursos de Word
Public Sub CerrarDocumento()
End Sub

' Lee el contenido de texto de un documento Word
' ParÃ¡metros:
'   - rutaDocumento: Ruta completa al archivo Word
' Retorna: El contenido de texto del documento o cadena vacÃ­a si falla
Public Function LeerContenidoDocumento(ByVal rutaDocumento As String) As String
End Function
