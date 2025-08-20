VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: CSolicitudRepository
' DESCRIPCI├ôN: Implementaci├│n real del repositorio de solicitudes PC
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

Option Explicit
Implements ISolicitudRepository

'******************************************************************************
' IMPLEMENTACI├ôN DE INTERFAZ ISolicitudRepository
'******************************************************************************

Private Function ISolicitudRepository_SavePC(solicitudPC As CSolicitudPC) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim solicitudId As Long
    
    ' Iniciar transacci├│n
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    ws.BeginTrans
    
    ' Verificar si la solicitud ya existe
    If solicitudPC.idSolicitud > 0 Then
        If SolicitudExists(solicitudPC.idSolicitud) Then
            ' Actualizar solicitud existente
            UpdateExistingSolicitud solicitudPC
            solicitudId = solicitudPC.idSolicitud
        Else
            ' Insertar nueva solicitud
            solicitudId = InsertNewSolicitud(solicitudPC)
        End If
    Else
        ' Insertar nueva solicitud
        solicitudId = InsertNewSolicitud(solicitudPC)
    End If
    
    ' Confirmar transacci├│n
    ws.CommitTrans
    ISolicitudRepository_SavePC = solicitudId
    
    Exit Function
    
ErrorHandler:
    ws.Rollback
    ISolicitudRepository_SavePC = 0
End Function

Private Function ISolicitudRepository_LoadPC(ByVal solicitudId As Long) As CSolicitudPC
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim solicitud As CSolicitudPC
    
    Set db = CurrentDb
    
    ' Consulta JOIN para obtener datos de ambas tablas
    sql = "SELECT s.*, pc.* FROM Tb_Solicitudes s " & _
          "INNER JOIN Tb_Datos_PC pc ON s.ID_Solicitud = pc.ID_Solicitud " & _
          "WHERE s.ID_Solicitud = " & solicitudId
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        Set solicitud = New CSolicitudPC
        
        ' Cargar propiedades de solicitud
        solicitud.idSolicitud = rs("ID_Solicitud")
        solicitud.IDExpediente = rs("ID_Expediente")
        solicitud.TipoSolicitud = rs("Tipo_Solicitud")
        solicitud.CodigoSolicitud = rs("Codigo_Solicitud")
        solicitud.EstadoInterno = rs("Estado_Interno")
        
        ' Cargar datos espec├¡ficos de PC
        Dim datosPC As T_Datos_PC
        datosPC.Procesador = rs("Procesador")
        datosPC.RAM = rs("RAM")
        datosPC.Almacenamiento = rs("Almacenamiento")
        datosPC.SistemaOperativo = rs("Sistema_Operativo")
        solicitud.DatosPC = datosPC
        
        Set ISolicitudRepository_LoadPC = solicitud
    Else
        Set ISolicitudRepository_LoadPC = Nothing
    End If
    
    rs.Close
    Exit Function
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    Set ISolicitudRepository_LoadPC = Nothing
End Function

'******************************************************************************
' M├ëTODOS AUXILIARES PRIVADOS
'******************************************************************************

Private Function SolicitudExists(ByVal ID As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    sql = "SELECT COUNT(*) as Total FROM Tb_Solicitudes WHERE ID_Solicitud = " & ID
    Set rs = db.OpenRecordset(sql)
    
    SolicitudExists = (rs("Total") > 0)
    rs.Close
End Function

Private Function InsertNewSolicitud(solicitud As CSolicitudPC) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim newId As Long
    
    Set db = CurrentDb
    
    ' Insertar en Tb_Solicitudes
    Set rs = db.OpenRecordset("Tb_Solicitudes", dbOpenDynaset)
    rs.AddNew
    rs("ID_Expediente") = solicitud.IDExpediente
    rs("Tipo_Solicitud") = solicitud.TipoSolicitud
    rs("Codigo_Solicitud") = solicitud.CodigoSolicitud
    rs("Estado_Interno") = solicitud.EstadoInterno
    rs.Update
    rs.Bookmark = rs.LastModified
    newId = rs("ID_Solicitud")
    rs.Close
    
    ' Insertar en Tb_Datos_PC
    Set rs = db.OpenRecordset("Tb_Datos_PC", dbOpenDynaset)
    rs.AddNew
    rs("ID_Solicitud") = newId
    rs("Procesador") = solicitud.DatosPC.Procesador
    rs("RAM") = solicitud.DatosPC.RAM
    rs("Almacenamiento") = solicitud.DatosPC.Almacenamiento
    rs("Sistema_Operativo") = solicitud.DatosPC.SistemaOperativo
    rs.Update
    rs.Close
    
    InsertNewSolicitud = newId
End Function

Private Sub UpdateExistingSolicitud(solicitud As CSolicitudPC)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Actualizar Tb_Solicitudes
    sql = "SELECT * FROM Tb_Solicitudes WHERE ID_Solicitud = " & solicitud.idSolicitud
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs("ID_Expediente") = solicitud.IDExpediente
        rs("Tipo_Solicitud") = solicitud.TipoSolicitud
        rs("Codigo_Solicitud") = solicitud.CodigoSolicitud
        rs("Estado_Interno") = solicitud.EstadoInterno
        rs.Update
    End If
    rs.Close
    
    ' Actualizar Tb_Datos_PC
    sql = "SELECT * FROM Tb_Datos_PC WHERE ID_Solicitud = " & solicitud.idSolicitud
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs("Procesador") = solicitud.DatosPC.Procesador
        rs("RAM") = solicitud.DatosPC.RAM
        rs("Almacenamiento") = solicitud.DatosPC.Almacenamiento
        rs("Sistema_Operativo") = solicitud.DatosPC.SistemaOperativo
        rs.Update
    End If
    rs.Close
End Sub
