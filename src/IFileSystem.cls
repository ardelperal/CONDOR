Option Compare Database
Option Explicit

' =====================================================
' INTERFAZ: IFileSystem
' DESCRIPCION: Contrato para interactuar con el sistema de ficheros
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - AbstracciÃ³n del FileSystemObject
' =====================================================

' Contrato para operaciones de sistema de ficheros
' Permite abstraer la interacciÃ³n con archivos para facilitar las pruebas unitarias

' Abre un archivo de texto para lectura/escritura
' @param Path: Ruta del archivo
' @param Mode: Modo de apertura (1=ForReading, 2=ForWriting, 8=ForAppending)
' @param Create: Si debe crear el archivo si no existe
' @return: Objeto TextStream para manipular el archivo
Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
End Function
