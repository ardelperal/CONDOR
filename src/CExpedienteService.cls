VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

Implements IExpedienteService

' ============================================================================
' Clase: CExpedienteService
' Descripción: Implementación real del servicio de expedientes
'              Conecta con la aplicación de Expedientes existente
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Implementación de la interfaz IExpedienteService
' Obtiene un expediente por su ID desde la base de datos de expedientes
Private Function IExpedienteService_GetExpedientePorID(ByVal IDExpediente As Long) As T_Expediente
    Dim resultado As T_Expediente
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    ' Inicializar resultado vacío
    resultado.IDExpediente = 0
    resultado.Nemotecnico = ""
    resultado.Titulo = ""
    resultado.ResponsableCalidad = ""
    resultado.ResponsableTecnico = ""
    resultado.Pecal = ""
    
    ' Obtener la ruta de la base de datos de expedientes desde configuración
    Dim rutaExpedientes As String
    rutaExpedientes = GetExpedientesDbPath()
    
    If rutaExpedientes = "" Then
        Debug.Print "Error: No se ha configurado la ruta de la base de datos de expedientes"
        GoTo Cleanup
    End If
    
    ' Abrir conexión a la base de datos de expedientes
    Set db = DBEngine.OpenDatabase(rutaExpedientes)
    
    ' Construir la consulta SQL con filtro por ID
    sql = "SELECT TbExpedientes.IDExpediente, TbExpedientes.Nemotecnico, TbExpedientes.Titulo, " & _
          "TbUsuariosAplicaciones_1.Nombre AS ResponsableCalidad, " & _
          "TbUsuariosAplicaciones.Nombre AS ResponsableTecnico, " & _
          "TbExpedientes.Pecal " & _
          "FROM ((TbExpedientes INNER JOIN TbExpedientesResponsables ON TbExpedientes.IDExpediente = TbExpedientesResponsables.IdExpediente) " & _
          "LEFT JOIN TbUsuariosAplicaciones ON TbExpedientesResponsables.IdUsuario = TbUsuariosAplicaciones.Id) " & _
          "LEFT JOIN TbUsuariosAplicaciones AS TbUsuariosAplicaciones_1 ON TbExpedientes.IDResponsableCalidad = TbUsuariosAplicaciones_1.Id " & _
          "WHERE (((TbExpedientesResponsables.EsJefeProyecto)='Sí') AND ((TbExpedientes.IDExpediente)=" & IDExpediente & "));"
    
    ' Ejecutar la consulta
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Si encontramos un resultado, llenar la estructura
    If Not rs.EOF Then
        resultado.IDExpediente = rs("IDExpediente")
        resultado.Nemotecnico = Nz(rs("Nemotecnico"), "")
        resultado.Titulo = Nz(rs("Titulo"), "")
        resultado.ResponsableCalidad = Nz(rs("ResponsableCalidad"), "")
        resultado.ResponsableTecnico = Nz(rs("ResponsableTecnico"), "")
        resultado.Pecal = Nz(rs("Pecal"), "")
    End If
    
Cleanup:
    ' Limpiar recursos
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    
    IExpedienteService_GetExpedientePorID = resultado
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en CExpedienteService.GetExpedientePorID: " & Err.Description
    Resume Cleanup
End Function

#End If

