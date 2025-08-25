Option Compare Database
Option Explicit
' CNotificationService.cls
' ImplementaciÃ³n segura del servicio de notificaciones asÃ­ncrono
' Reconstruido para cumplir con principios de inyecciÃ³n de dependencias y seguridad
' Basado en las Especificaciones de IntegraciÃ³n - SecciÃ³n 2

' Implementa la interfaz INotificationService
Implements INotificationService

' Variables privadas para dependencias inyectadas
Private m_Config As IConfig
Private m_OperationLogger As IOperationLogger

' MÃ©todo pÃºblico para inyectar dependencias
Public Sub Initialize(config As IConfig, logger As IOperationLogger)
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Set m_Config = config
    Set m_OperationLogger = logger
    
    Exit Sub
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CNotificationService.Initialize"
End Sub

' ImplementaciÃ³n del mÃ©todo EnviarNotificacion de la interfaz
Private Function INotificationService_EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    Dim errorHandler As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    ' Validar que las dependencias estÃ©n inyectadas
    If m_Config Is Nothing Or m_OperationLogger Is Nothing Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Validar parÃ¡metros
    If Len(Trim(destinatarios)) = 0 Or Len(Trim(asunto)) = 0 Or Len(Trim(cuerpoHTML)) = 0 Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Obtener configuraciÃ³n desde IConfig inyectado
    Dim rutaCorreosDB As String
    Dim passwordDB As String
    Dim usuarioActual As String
    Dim correoAdministrador As String
    
    rutaCorreosDB = m_Config.GetCorreosDBPath()
    passwordDB = m_Config.GetDatabasePassword()
    usuarioActual = m_Config.GetUsuarioActual()
    correoAdministrador = m_Config.GetCorreoAdministrador()
    
    If Len(rutaCorreosDB) = 0 Then
        INotificationService_EnviarNotificacion = False
        Exit Function
    End If
    
    ' Conectar a la base de datos de correos
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase(rutaCorreosDB, False, False, ";PWD=" & passwordDB)
    
    ' Calcular nuevo IDCorreo
    Dim nuevoID As Long
    nuevoID = Nz(DMax("IDCorreo", "TbCorreosEnviados", "", db), 0) + 1
    
    ' Crear consulta parametrizada para eliminar vulnerabilidad SQL injection
    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("", _
        "INSERT INTO TbCorreosEnviados (IDCorreo, Aplicacion, Asunto, Cuerpo, " & _
        "Destinatarios, DestinatariosConCopia, DestinatariosConCopiaOculta, " & _
        "URLAdjunto, FechaGrabacion, FechaEnvio) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
    
    ' Asignar parÃ¡metros de forma segura
    qdf.Parameters(0) = nuevoID
    qdf.Parameters(1) = "CONDOR"
    qdf.Parameters(2) = asunto
    qdf.Parameters(3) = cuerpoHTML
    qdf.Parameters(4) = destinatarios
    qdf.Parameters(5) = usuarioActual
    qdf.Parameters(6) = correoAdministrador
    qdf.Parameters(7) = urlAdjunto
    qdf.Parameters(8) = Now()
    qdf.Parameters(9) = Null
    
    ' Ejecutar consulta parametrizada
    qdf.Execute
    
    ' Limpiar recursos
    qdf.Close
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Registrar operaciÃ³n exitosa
    m_OperationLogger.LogOperation "NOTIFICATION", CStr(nuevoID), _
        "Correo encolado exitosamente. ID: " & nuevoID & ", Destinatarios: " & destinatarios
    
    INotificationService_EnviarNotificacion = True
    Exit Function
    
ErrorHandler:
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "CNotificationService.EnviarNotificacion"
    
    ' Limpiar recursos en caso de error
    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    
    INotificationService_EnviarNotificacion = False
End Function

' Destructor
Private Sub Class_Terminate()
    Set m_Config = Nothing
    Set m_OperationLogger = Nothing
End Sub













