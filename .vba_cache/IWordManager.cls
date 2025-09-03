VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ============================================================================
' Interfaz: IWordManager
' Descripción: Define el contrato para el manejo de documentos Word.
' Autor: CONDOR-Expert
' Fecha: 2025-08-22
' Versión: 1.0
' ============================================================================

' ============================================================================
' MÉTODOS PARA MANEJO DE WORD
' ============================================================================

' Abre un documento Word desde una ruta específica
' Parámetros:
'   - rutaDocumento: Ruta completa al archivo Word
' Retorna: True si se abrió correctamente, False en caso contrario
Public Function AbrirDocumento(ByVal rutaDocumento As String) As Boolean
End Function

' Reemplaza texto en el documento actualmente abierto
' Parámetros:
'   - textoABuscar: Texto que se va a reemplazar
'   - textoReemplazo: Texto de reemplazo
' Retorna: True si se realizó el reemplazo, False en caso contrario
Public Function ReemplazarTexto(ByVal textoABuscar As String, ByVal textoReemplazo As String) As Boolean
End Function

' Guarda el documento actualmente abierto en una ruta específica
' Parámetros:
'   - rutaDestino: Ruta donde guardar el documento
' Retorna: True si se guardó correctamente, False en caso contrario
Public Function GuardarDocumento(ByVal rutaDestino As String) As Boolean
End Function

' Cierra el documento y libera los recursos de Word
Public Sub CerrarDocumento()
End Sub

' Lee el contenido de texto de un documento Word
' Parámetros:
'   - rutaDocumento: Ruta completa al archivo Word
' Retorna: El contenido de texto del documento o cadena vacía si falla
Public Function LeerContenidoDocumento(ByVal rutaDocumento As String) As String
End Function




