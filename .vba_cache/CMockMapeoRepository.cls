VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockMapeoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IMapeoRepository

' --- Variable de estado para configurar el comportamiento ---
Private m_GetMapeoPorTipo_Result As DAO.Recordset

' --- Método público para configurar el mock desde los tests ---
Public Sub ConfigureGetMapeoPorTipo(ByVal rs As DAO.Recordset)
    Set m_GetMapeoPorTipo_Result = rs
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    Set m_GetMapeoPorTipo_Result = Nothing
End Sub

' --- Implementación de la Interfaz ---
Private Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.Recordset
    ' Devuelve el recordset configurado por el test
    Set IMapeoRepository_GetMapeoPorTipo = m_GetMapeoPorTipo_Result
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================

' Método público de conveniencia para GetMapeoPorTipo
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As DAO.Recordset
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(tipoSolicitud)
End Function