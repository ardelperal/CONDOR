Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IAuthRepository

' Variables privadas para almacenar los datos mock
Private m_MockUserRecordset As DAO.Recordset
Private m_MockPermissionsRecordset As DAO.Recordset

'******************************************************************************
' MÉTODO: SetMockUserData
' DESCRIPCIÓN: Configura el Recordset que debe devolver GetUserByEmail
' PARÁMETROS: mockRecordset - Recordset configurado para las pruebas
'******************************************************************************
Public Sub SetMockUserData(mockRecordset As DAO.Recordset)
    Set m_MockUserRecordset = mockRecordset
End Sub

'******************************************************************************
' MÉTODO: SetMockPermissionsData
' DESCRIPCIÓN: Configura el Recordset que debe devolver GetUserApplicationPermissions
' PARÁMETROS: mockRecordset - Recordset configurado para las pruebas
'******************************************************************************
Public Sub SetMockPermissionsData(mockRecordset As DAO.Recordset)
    Set m_MockPermissionsRecordset = mockRecordset
End Sub

'******************************************************************************
' IMPLEMENTACIÓN DE IAuthRepository
'******************************************************************************

Private Function IAuthRepository_GetUserByEmail(ByVal UserEmail As String) As DAO.Recordset
    ' El mock devuelve el Recordset configurado mediante SetMockUserData
    ' independientemente del email proporcionado
    Set IAuthRepository_GetUserByEmail = m_MockUserRecordset
End Function

Private Function IAuthRepository_GetUserApplicationPermissions(ByVal UserEmail As String, ByVal ApplicationName As String) As DAO.Recordset
    ' El mock devuelve el Recordset configurado mediante SetMockPermissionsData
    ' independientemente de los parámetros proporcionados
    Set IAuthRepository_GetUserApplicationPermissions = m_MockPermissionsRecordset
End Function

'******************************************************************************
' MÉTODOS PÚBLICOS DE CONVENIENCIA
'******************************************************************************

Public Function GetUserByEmail(ByVal UserEmail As String) As DAO.Recordset
    Set GetUserByEmail = IAuthRepository_GetUserByEmail(UserEmail)
End Function

Public Function GetUserApplicationPermissions(ByVal UserEmail As String, ByVal ApplicationName As String) As DAO.Recordset
    Set GetUserApplicationPermissions = IAuthRepository_GetUserApplicationPermissions(UserEmail, ApplicationName)
End Function

#End If