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

' Establece el texto de un marcador (bookmark) en el documento
' Parámetros:
'   - BookmarkName: Nombre del marcador
'   - Value: Valor a establecer en el marcador
' Retorna: True si se estableció correctamente, False en caso contrario
Public Function SetBookmarkText(ByVal BookmarkName As String, ByVal Value As String) As Boolean
End Function

' Obtiene el texto de un marcador (bookmark) del documento
' Parámetros:
'   - BookmarkName: Nombre del marcador
' Retorna: El texto del marcador o cadena vacía si no existe
Public Function GetBookmarkText(ByVal BookmarkName As String) As String
End Function

' Guarda el documento actualmente abierto
' Retorna: True si se guardó correctamente, False en caso contrario
Public Function GuardarDocumento() As Boolean
End Function
Public Sub Dispose()
End Sub



