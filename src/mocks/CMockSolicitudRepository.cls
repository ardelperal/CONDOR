VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockSolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: CMockSolicitudRepository
' DESCRIPCIÓN: Mock del repositorio de solicitudes PC para pruebas
' AUTOR: Sistema CONDOR
' FECHA: 2024
'******************************************************************************

#If DEV_MODE Then

Option Explicit
Implements ISolicitudRepository

' Almacenamiento en memoria para las pruebas
Private mSolicitudes As Collection
Private mNextId As Long

'******************************************************************************
' CONSTRUCTOR Y DESTRUCTOR
'******************************************************************************

Private Sub Class_Initialize()
    Set mSolicitudes = New Collection
    mNextId = 1
    
    ' Precargar algunos datos de prueba
    LoadTestData
End Sub

Private Sub Class_Terminate()
    Set mSolicitudes = Nothing
End Sub

'******************************************************************************
' IMPLEMENTACIÓN DE INTERFAZ ISolicitudRepository
'******************************************************************************

Private Function ISolicitudRepository_SavePC(solicitudPC As CSolicitudPC) As Long
    Dim solicitudClone As CSolicitudPC
    Dim keyToRemove As String
    
    ' Clonar la solicitud para evitar referencias
    Set solicitudClone = CloneSolicitud(solicitudPC)
    
    ' Si no tiene ID, asignar uno nuevo
    If solicitudClone.idSolicitud = 0 Then
        solicitudClone.idSolicitud = mNextId
        mNextId = mNextId + 1
    Else
        ' Si ya existe, eliminar la versión anterior
        keyToRemove = "ID_" & solicitudClone.idSolicitud
        On Error Resume Next
        mSolicitudes.Remove keyToRemove
        On Error GoTo 0
    End If
    
    ' Agregar a la colección
    mSolicitudes.Add solicitudClone, "ID_" & solicitudClone.idSolicitud
    
    ISolicitudRepository_SavePC = solicitudClone.idSolicitud
End Function

Private Function ISolicitudRepository_LoadPC(ByVal solicitudId As Long) As CSolicitudPC
    Dim key As String
    Dim solicitud As CSolicitudPC
    
    key = "ID_" & solicitudId
    
    On Error Resume Next
    Set solicitud = mSolicitudes(key)
    On Error GoTo 0
    
    If Not solicitud Is Nothing Then
        ' Retornar una copia para evitar modificaciones accidentales
        Set ISolicitudRepository_LoadPC = CloneSolicitud(solicitud)
    Else
        Set ISolicitudRepository_LoadPC = Nothing
    End If
End Function

'******************************************************************************
' MÉTODOS AUXILIARES PARA PRUEBAS
'******************************************************************************

Private Function CloneSolicitud(original As CSolicitudPC) As CSolicitudPC
    Dim clon As CSolicitudPC
    Set clon = New CSolicitudPC
    
    ' Copiar propiedades básicas
    clon.idSolicitud = original.idSolicitud
    clon.IDExpediente = original.IDExpediente
    clon.TipoSolicitud = original.TipoSolicitud
    clon.CodigoSolicitud = original.CodigoSolicitud
    clon.EstadoInterno = original.EstadoInterno
    clon.DatosPC = original.DatosPC
    
    Set CloneSolicitud = clon
End Function

Public Sub ClearData()
    Set mSolicitudes = New Collection
    mNextId = 1
End Sub

Public Function ExistsSolicitud(ByVal ID As Long) As Boolean
    Dim key As String
    Dim solicitud As CSolicitudPC
    
    key = "ID_" & ID
    
    On Error Resume Next
    Set solicitud = mSolicitudes(key)
    On Error GoTo 0
    
    ExistsSolicitud = Not (solicitud Is Nothing)
End Function

Private Sub LoadTestData()
    ' Datos de prueba predefinidos
    Dim solicitud As CSolicitudPC
    Dim datosPC As T_Datos_PC
    
    ' Solicitud de prueba 1
    Set solicitud = New CSolicitudPC
    solicitud.idSolicitud = 999
    solicitud.IDExpediente = 1001
    solicitud.TipoSolicitud = "PC"
    solicitud.CodigoSolicitud = "PC-TEST-001"
    solicitud.EstadoInterno = "Pendiente"
    
    datosPC.Procesador = "Intel i7-12700K"
    datosPC.RAM = "32GB DDR4"
    datosPC.Almacenamiento = "1TB NVMe SSD"
    datosPC.SistemaOperativo = "Windows 11 Pro"
    solicitud.DatosPC = datosPC
    
    mSolicitudes.Add solicitud, "ID_999"
    mNextId = 1000
End Sub

#End If