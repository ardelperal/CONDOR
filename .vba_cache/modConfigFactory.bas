Attribute VB_Name = "modConfigFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modConfigFactory
' DESCRIPCIÓN: Factory para crear el servicio de configuración.
'              Implementa la lógica de bootstrap para cargar
'              la configuración desde el backend correcto.
' AUTOR: CONDOR-Architect
' FECHA: 2025-08-28
' =====================================================

' --- CONSTANTES DE RUTA BASE ---
' Usadas para construir la ruta al backend según el entorno
Private Const DEV_BACKEND_PATH As String = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
Private Const PROD_BACKEND_PATH As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\CONDOR\CONDOR_datos.accdb"

' --- FUNCIÓN FACTORY PRINCIPAL ---
Public Function CreateConfigService() As IConfig
    On Error GoTo ErrorHandler
    
    Dim configImpl As New CConfig
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendPath As String
    
    ' 1. Determinar la ruta del Backend leyendo TbLocalConfig del Frontend
    Set db = CurrentDb ' Correcto aquí, lee la tabla local del frontend
    Set rs = db.OpenRecordset("SELECT Entorno FROM TbLocalConfig")
    
    If Not rs.EOF Then
        Select Case UCase(Trim(rs("Entorno")))
            Case "LOCAL"
                backendPath = DEV_BACKEND_PATH
            Case "OFICINA"
                backendPath = PROD_BACKEND_PATH
            Case Else
                ' Error: Entorno no reconocido
                Err.Raise vbObjectError + 1001, "modConfigFactory", "Entorno no válido en TbLocalConfig: " & rs("Entorno")
        End Select
    Else
        ' Error: TbLocalConfig está vacía
        Err.Raise vbObjectError + 1002, "modConfigFactory", "La tabla TbLocalConfig está vacía o no se encontró."
    End If
    
    rs.Close
    db.Close
    
    ' 2. Abrir el Backend y cargar la configuración en el objeto CConfig
    Set db = DBEngine.OpenDatabase(backendPath, False, True) ' Abrir en solo lectura
    configImpl.Load db
    db.Close
    
    ' 3. Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateConfigService = configImpl
    
    Exit Function
    
ErrorHandler:
    ' En caso de error en la factoría, no podemos usar el ErrorHandler.
    ' Mostramos un mensaje crítico y devolvemos Nothing.
    MsgBox "Error CRÍTICO al cargar la configuración de la aplicación." & vbCrLf & vbCrLf & _
           "Módulo: modConfigFactory" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical, "Fallo de Arranque de CONDOR"
    Set CreateConfigService = Nothing
End Function