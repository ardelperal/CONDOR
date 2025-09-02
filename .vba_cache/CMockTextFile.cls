VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' =====================================================
' CLASE: CMockTextFile
' DESCRIPCION: Mock del objeto TextFile para pruebas unitarias
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Mock para aislamiento de pruebas
' =====================================================

' Mock del objeto TextFile que simula las operaciones
' de escritura en archivos sin interactuar con el disco

' Propiedades públicas para verificación en pruebas
Public WasWriteLineCalled As Boolean
Public WasCloseCalled As Boolean
Public LastWrittenLine As String
Public WriteLineCallCount As Long
Public CloseCallCount As Long
Public AllWrittenLines As Collection

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
    Set AllWrittenLines = Nothing
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS PARA CONFIGURACIÓN DEL MOCK
' ============================================================================

' Reinicia el estado del mock para una nueva prueba
Public Sub Reset()
    WasWriteLineCalled = False
    WasCloseCalled = False
    LastWrittenLine = ""
    WriteLineCallCount = 0
    CloseCallCount = 0
    Set AllWrittenLines = New Collection
End Sub

' ============================================================================
' MÉTODOS QUE SIMULAN EL COMPORTAMIENTO DEL TEXTFILE REAL
' ============================================================================

' Simula la escritura de una línea en el archivo
Public Sub WriteLine(ByVal line As String)
    WasWriteLineCalled = True
    LastWrittenLine = line
    WriteLineCallCount = WriteLineCallCount + 1
    AllWrittenLines.Add line
End Sub

' Simula el cierre del archivo
Public Sub Close()
    WasCloseCalled = True
    CloseCallCount = CloseCallCount + 1
End Sub




