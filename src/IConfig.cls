VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Interfaz: IConfig
' Descripci?n: Define el contrato para servicios de configuraci?n del sistema CONDOR
' Arquitectura: Capa de Servicios - Interface

' Obtiene un valor de configuraci?n por su clave
' @param key: Clave de configuraci?n a buscar
' @return: Valor de configuraci?n (Variant para flexibilidad de tipos)
Public Function GetValue(ByVal key As String) As Variant
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Function

' Establece un valor de configuraci?n
' @param key: Clave de configuraci?n
' @param value: Valor a establecer
Public Sub SetValue(ByVal key As String, ByVal value As Variant)
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Sub

' Inicializa el entorno de configuraci?n
' @return: True si la inicializaci?n fue exitosa
Public Function InitializeEnvironment() As Boolean
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Function

' Verifica si una clave de configuraci?n existe
' @param key: Clave a verificar
' @return: True si la clave existe
Public Function HasKey(ByVal key As String) As Boolean
    ' Esta es una interfaz - debe ser implementada por las clases concretas
End Function




