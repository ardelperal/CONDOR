VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CNotificationService.cls
' Implementación del servicio de notificaciones asíncrono
' Basado en las Especificaciones de Integración - Sección 2
' Encola correos en la base de datos Correos_datos.accdb

Option Compare Database
Option Explicit

' Implementa la interfaz INotificationService
Implements INotificationService

' Variables privadas
Private m_configService As IConfigService

' Constructor
Private Sub Class_Initialize()
    ' Inicializar dependencias
    Set m_configService = New CConfigService
End Sub

' Destructor
Private Sub Class_Terminate()
    Set m_configService = Nothing
End Sub

' Implementación del método EnviarNotificacion de la interfaz
Private Function INotificationService_EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Validar parámetros
    If Len(Trim(destinatarios)) = 0 Or Len(Trim(asunto)) = 0 Or Len(Trim(cuerpoHTML)) = 0 Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Obtener ruta de la base de datos de correos
    Dim rutaCorreosDB As String
    rutaCorreosDB = m_configService.ObtenerRutaCorreosDB()
    
    If Len(rutaCorreosDB) = 0 Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Conectar a la base de datos de correos
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(rutaCorreosDB)
    
    ' Calcular nuevo IDCorreo
    Dim nuevoID As Long
    nuevoID = Nz(DMax("IDCorreo", "TbCorreosEnviados", "", db), 0) + 1
    
    ' Obtener usuario actual para DestinatariosConCopia
    Dim usuarioActual As String
    usuarioActual = m_configService.ObtenerUsuarioActual()
    
    ' Obtener correo del administrador para DestinatariosConCopiaOculta
    Dim correoAdministrador As String
    correoAdministrador = m_configService.ObtenerCorreoAdministrador()
    
    ' Construir y ejecutar INSERT
    Dim sql As String
    sql = "INSERT INTO TbCorreosEnviados (IDCorreo, Aplicacion, Asunto, Cuerpo, " & _
          "Destinatarios, DestinatariosConCopia, DestinatariosConCopiaOculta, " & _
          "URLAdjunto, FechaGrabacion, FechaEnvio) VALUES (" & _
          nuevoID & ", 'CONDOR', '" & Replace(asunto, "'", "''") & "', '" & _
          Replace(cuerpoHTML, "'", "''") & "', '" & Replace(destinatarios, "'", "''") & "', '" & _
          Replace(usuarioActual, "'", "''") & "', '" & Replace(correoAdministrador, "'", "''") & "', '" & _
          Replace(urlAdjunto, "'", "''") & "', Now(), Null)"
    
    db.Execute sql
    
    ' Cerrar conexión
    db.Close
    Set db = Nothing
    
    INotificationService_EnviarNotificacion = True
    Exit Function
    
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    INotificationService_EnviarNotificacion = False
End Function