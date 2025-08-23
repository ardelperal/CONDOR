Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit
' Factory para el servicio de configuración. Versión 4.0.

' Constante global para la ruta de la base de datos del backend
Public Const BACKEND_DB_PATH As String = "\..\data\CONDOR_Backend.accdb"

Public Function CreateConfigService() As IConfig
    Dim config As New CConfig
    ' Aquí se podría inicializar si fuera necesario
    Set CreateConfigService = config
End Function











