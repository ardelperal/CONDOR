VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IAuthRepository

Private m_MockData As EAuthData

' Método para que los tests configuren los datos que debe devolver el mock
Public Sub ConfigureMockData(ByVal data As EAuthData)
    Set m_MockData = data
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    Set m_MockData = Nothing
End Sub

Private Function IAuthRepository_GetUserAuthData(ByVal userEmail As String) As EAuthData
    ' Si se han configurado datos, devolverlos. Si no, devolver el valor por defecto.
    If Not m_MockData Is Nothing Then
        Set IAuthRepository_GetUserAuthData = m_MockData
    Else
        Dim defaultData As New EAuthData
        defaultData.UserExists = False
        Set IAuthRepository_GetUserAuthData = defaultData
    End If
End Function