Attribute VB_Name = "modErrorHandler"
' ===============================================================================
' M├│dulo: modErrorHandler
' Descripci├│n: Sistema centralizado de manejo de errores para CONDOR
' Autor: Sistema CONDOR
' Fecha: 2024
' ===============================================================================

Option Compare Database
Option Explicit

' ===============================================================================
' FUNCI├ôN PRINCIPAL DE REGISTRO DE ERRORES
' ===============================================================================

' Registra un error en la base de datos y opcionalmente notifica al administrador
' @param errNumber: N├║mero del error
' @param errDescription: Descripci├│n del error
' @param errSource: Origen del error (Clase.Funci├│n)
' @param userAction: Acci├│n que estaba realizando el usuario (opcional)
Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, ByVal errSource As String, Optional ByVal userAction As String = "")
    On Error Resume Next ' Para evitar bucles infinitos si el propio logging falla

    Dim logger As ILoggingService
    Set logger = New CLoggingService ' TODO: Usar Factory cuando exista

    ' Registrar error en el servicio centralizado con contexto
    Dim contexto As String
    contexto = "AccionUsuario=" & userAction & "; Usuario=" & Environ("USERNAME")
    logger.LogError errNumber, errDescription, errSource, contexto

    ' Notificar si es cr├¡tico (sin depender de base de datos)
    If IsCriticalError(errNumber) Then
        Call CreateAdminNotification(errNumber, errDescription, errSource, Environ("USERNAME"))
    End If

    Set logger = Nothing
End Sub

' ===============================================================================
' FUNCIONES DE APOYO
' ===============================================================================

' Obtiene la ruta de la base de datos de datos
Private Function GetDatabasePath() As String
    GetDatabasePath = CurrentProject.Path & "\CONDOR_datos.accdb"
End Function

' Determina si un error es cr├¡tico y requiere notificaci├│n al administrador
Private Function IsCriticalError(ByVal errNumber As Long) As Boolean
    Select Case errNumber
        Case 3024, 3044, 3051, 3078, 3343 ' Errores de base de datos cr├¡ticos
            IsCriticalError = True
        Case 7, 9, 11, 13 ' Errores de memoria y tipos cr├¡ticos
            IsCriticalError = True
        Case 3265, 3421, 3709 ' Errores de conexi├│n y permisos
            IsCriticalError = True
        Case Else
            IsCriticalError = False
    End Select
End Function

' Crea una notificaci├│n para el administrador en caso de error cr├¡tico
Private Sub CreateAdminNotification(errNumber As Long, errDescription As String, errSource As String, usuario As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim strSQL As String
    Dim strAsunto As String
    Dim strMensaje As String
    
    ' Preparar el mensaje de notificaci├│n
    strAsunto = "ERROR CR├ìTICO en CONDOR - " & errSource
    strMensaje = "Se ha producido un error cr├¡tico en el sistema CONDOR:" & vbCrLf & vbCrLf & _
                 "Error: " & errNumber & " - " & errDescription & vbCrLf & _
                 "Origen: " & errSource & vbCrLf & _
                 "Usuario: " & usuario & vbCrLf & _
                 "Fecha/Hora: " & Format(Now(), "dd/mm/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
                 "Por favor, revise el sistema lo antes posible."
    
    ' Conectar a la base de datos
    Set db = OpenDatabase(GetDatabasePath())
    
    ' Insertar en la cola de correos (si existe la tabla)
    strSQL = "INSERT INTO Tb_Cola_Correos (" & _
             "Destinatario, " & _
             "Asunto, " & _
             "Mensaje, " & _
             "Fecha_Creacion, " & _
             "Estado, " & _
             "Prioridad" & _
             ") VALUES (" & _
             "'administrador@condor.local', " & _
             "'" & Replace(strAsunto, "'", "''") & "', " & _
             "'" & Replace(strMensaje, "'", "''") & "', " & _
             "'" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "', " & _
             "'Pendiente', " & _
             "'Alta'" & _
             ")"
    
    db.Execute strSQL
    
    ' Limpiar recursos
    db.Close
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' Si no se puede crear la notificaci├│n, registrar en log local
    Call WriteToLocalLog("ERROR al crear notificaci├│n admin: " & Err.Description)
    
    ' Limpiar recursos
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Sub

' Escribe en un archivo de log local como ├║ltimo recurso
Private Sub WriteToLocalLog(mensaje As String)
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logPath As String
    
    logPath = CurrentProject.Path & "\condor_error.log"
    fileNum = FreeFile
    
    Open logPath For Append As #fileNum
    Print #fileNum, Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & mensaje
    Close #fileNum
End Sub

' ===============================================================================
' FUNCIONES P├ÜBLICAS DE UTILIDAD
' ===============================================================================

' Funci├│n de conveniencia para registrar errores desde bloques de manejo de errores
Public Sub LogCurrentError(errSource As String, Optional userAction As String = "")
    Call LogError(Err.Number, Err.Description, errSource, userAction)
End Sub

' Limpia logs antiguos (mantiene solo los ├║ltimos 30 d├¡as)
Public Sub CleanOldLogs()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim strSQL As String
    Dim fechaLimite As String
    
    fechaLimite = Format(DateAdd("d", -30, Date), "yyyy-mm-dd")
    
    Set db = OpenDatabase(GetDatabasePath())
    
    strSQL = "DELETE FROM Tb_Log_Errores WHERE Fecha_Hora < '" & fechaLimite & "'"
    db.Execute strSQL
    
    db.Close
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    Call WriteToLocalLog("ERROR en CleanOldLogs: " & Err.Description)
    
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
End Sub
