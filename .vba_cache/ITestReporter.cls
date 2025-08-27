VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Interfaz: ITestReporter
' Propósito: Define el contrato para clases que generan reportes de pruebas
' Arquitectura: Patrón Strategy para diferentes tipos de reportes

' Inicializa el reportero con los resultados de las pruebas
Public Sub Initialize(ByVal results As Collection)
End Sub

' Genera y devuelve el reporte como string
Public Function GenerateReport() As String
End Function

