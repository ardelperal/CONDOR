VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' Proyecto: CONDOR
' Clase: CMockAuthRepository
' Descripción: Implementación mock del repositorio de autenticación para pruebas
' Implementa: IAuthRepository
' Autor: Sistema CONDOR
' Fecha: 2024
' =============================================================================

Option Explicit

Implements IAuthRepository

' Variables privadas
Private m_MockData As T_AuthData

' Constructor
Private Sub Class_Initialize()
    Set m_MockData = New T_AuthData
End Sub

' Destructor
Private Sub Class_Terminate()
    Set m_MockData = Nothing
End Sub

' Configura los datos mock que devolverá el repositorio
' @param data: Objeto T_AuthData con los datos de prueba
Public Sub ConfigureMockData(data As T_AuthData)
    On Error GoTo ErrorHandler
    
    If data Is Nothing Then
        Err.Raise 5, "CMockAuthRepository.ConfigureMockData", "El parámetro data no puede ser Nothing"
    End If
    
    Set m_MockData = data
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "CMockAuthRepository.ConfigureMockData", Err.Description
End Sub

' Implementación de IAuthRepository_GetUserAuthData
' @param UserEmail: Email del usuario (no utilizado en el mock)
' @return: Objeto T_AuthData configurado previamente
Private Function IAuthRepository_GetUserAuthData(ByVal UserEmail As String) As T_AuthData
    On Error GoTo ErrorHandler
    
    If m_MockData Is Nothing Then
        Err.Raise 91, "CMockAuthRepository.IAuthRepository_GetUserAuthData", "No se han configurado datos mock. Llame a ConfigureMockData primero."
    End If
    
    Set IAuthRepository_GetUserAuthData = m_MockData
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "CMockAuthRepository.IAuthRepository_GetUserAuthData", Err.Description
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA - ESTÁNDAR DE ORO PARA MOCKS
' ============================================================================

' Método público de conveniencia para GetUserAuthData
Public Function GetUserAuthData(ByVal UserEmail As String) As T_AuthData
    Set GetUserAuthData = IAuthRepository_GetUserAuthData(UserEmail)
End Function

' Método para resetear el estado del mock
Public Sub Reset()
    Set m_MockData = New T_AuthData
End Sub

