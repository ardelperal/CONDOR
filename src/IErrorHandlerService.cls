VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorHandlerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' Interfaz: IErrorHandlerService
' Descripción: Define el contrato para servicios de manejo de errores del sistema CONDOR
' Arquitectura: Capa de Servicios - Interface

' Inicializa el servicio con sus dependencias
' @param config: Configuración del sistema
' @param fileSystem: Sistema de archivos
Public Sub Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
End Sub

' Registra un error en el sistema de logging
' @param errNumber: Número del error
' @param errDescription: Descripción del error
' @param moduleName: Nombre del módulo donde ocurrió el error (opcional)
' @param isCritical: Indica si el error es crítico (opcional, por defecto False)
Public Sub LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
End Sub

Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String)
End Sub

Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String)
End Sub












