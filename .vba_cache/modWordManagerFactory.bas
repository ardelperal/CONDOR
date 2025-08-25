Attribute VB_Name = "modWordManagerFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modWordManagerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de gestión de Word
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-15
' =====================================================

Public Function CreateWordManager() As IWordManager
    On Error GoTo ErrorHandler
    
    ' Crear una instancia de la clase concreta
    Dim wordManagerInstance As New CWordManager
    
    ' Devolver la instancia como el tipo de la interfaz
    Set CreateWordManager = wordManagerInstance
    
    Exit Function
    
ErrorHandler:
    Dim errorHandler As IErrorHandlerService
    Set errorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    errorHandler.LogError Err.Number, Err.Description, "modWordManagerFactory.CreateWordManager"
    Set CreateWordManager = Nothing
End Function