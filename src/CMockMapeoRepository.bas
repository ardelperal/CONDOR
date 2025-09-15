Option Compare Database
Option Explicit

Implements IMapeoRepository

' --- Variable de estado para configurar el comportamiento ---
Private m_GetMapeoPorTipo_Result As EMapeo
Private m_ObtenerMapeosPorCategoria_Result As Object

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' --- Método público para configurar el mock desde los tests ---
Public Sub ConfigureGetMapeoPorTipo(ByVal mapeo As EMapeo)
    Set m_GetMapeoPorTipo_Result = mapeo
End Sub

Public Sub ConfigureObtenerMapeosPorCategoria(ByVal mapeos As Object)
    Set m_ObtenerMapeosPorCategoria_Result = mapeos
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    Set m_GetMapeoPorTipo_Result = Nothing
    Set m_ObtenerMapeosPorCategoria_Result = Nothing
End Sub

' --- Implementación de la Interfaz ---
Private Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    ' Devuelve el mapeo configurado por el test
    Set IMapeoRepository_GetMapeoPorTipo = m_GetMapeoPorTipo_Result
End Function

' ============================================================================
' MÉTODOS PÚBLICOS DE CONVENIENCIA
' ============================================================================

' Método público de conveniencia para GetMapeoPorTipo
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(tipoSolicitud)
End Function

Private Function IMapeoRepository_ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    Set IMapeoRepository_ObtenerMapeosPorCategoria = m_ObtenerMapeosPorCategoria_Result
End Function

Public Function ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    Set ObtenerMapeosPorCategoria = IMapeoRepository_ObtenerMapeosPorCategoria(categoria)
End Function