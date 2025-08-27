VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IFileSystem

' =====================================================
' CLASE: CMockFileSystem
' DESCRIPCION: Mock del sistema de ficheros para pruebas unitarias
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Mock para aislamiento de pruebas
' =====================================================

' Mock del sistema de ficheros que permite verificar
' las llamadas realizadas durante las pruebas unitarias
' sin interactuar con el sistema de ficheros real

' Propiedades públicas para verificación en pruebas
Public WasOpenTextFileCalled As Boolean
Public LastPath As String
Public LastMode As Long
Public LastCreate As Boolean
Public CallCount As Long

' Mock del objeto TextFile que se devuelve
Private m_mockTextFile As CMockTextFile

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    WasOpenTextFileCalled = False
    LastPath = ""
    LastMode = 0
    LastCreate = False
    CallCount = 0
    Set m_mockTextFile = New CMockTextFile
End Sub

Private Sub Class_Terminate()
    Set m_mockTextFile = Nothing
End Sub

' ============================================================================
' MÉTODOS PÚBLICOS PARA CONFIGURACIÓN DEL MOCK
' ============================================================================

' Reinicia el estado del mock para una nueva prueba
Public Sub Reset()
    WasOpenTextFileCalled = False
    LastPath = ""
    LastMode = 0
    LastCreate = False
    CallCount = 0
    m_mockTextFile.Reset
End Sub

' Obtiene el mock del TextFile para configuración adicional
Public Function GetMockTextFile() As CMockTextFile
    Set GetMockTextFile = m_mockTextFile
End Function

' ============================================================================
' IMPLEMENTACIÃ“N DE LA INTERFAZ IFileSystem
' ============================================================================

' Mock de la apertura de archivos de texto
Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    ' Registrar la llamada para verificación
    WasOpenTextFileCalled = True
    LastPath = Path
    LastMode = Mode
    LastCreate = Create
    CallCount = CallCount + 1
    
    ' Devolver el mock del TextFile
    Set IFileSystem_OpenTextFile = m_mockTextFile
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para OpenTextFile
Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set OpenTextFile = IFileSystem_OpenTextFile(Path, Mode, Create)
End Function


