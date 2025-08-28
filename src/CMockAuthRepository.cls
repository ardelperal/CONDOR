Attribute VB_Name = "CMockAuthRepository"
Option Compare Database
Option Explicit
Implements IAuthRepository

Private m_MockData As E_AuthData

' Método para que los tests configuren los datos que debe devolver el mock
Public Sub ConfigureMockData(ByVal data As E_AuthData)
    Set m_MockData = data
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    Set m_MockData = Nothing
End Sub

Private Function IAuthRepository_GetUserAuthData(ByVal userEmail As String) As E_AuthData
    ' Si se han configurado datos, devolverlos. Si no, devolver el valor por defecto.
    If Not m_MockData Is Nothing Then
        Set IAuthRepository_GetUserAuthData = m_MockData
    Else
        Dim defaultData As New E_AuthData
        defaultData.UserExists = False
        Set IAuthRepository_GetUserAuthData = defaultData
    End If
End Function