Attribute VB_Name = "modErrorHandlerFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modErrorHandlerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de errores
' PATRÓN: CERO ARGUMENTOS (Lección 37)
' =====================================================

Public Function CreateErrorHandlerService() As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim errorHandlerImpl As New CErrorHandlerService
    
    ' ¡CUIDADO! ErrorHandler depende de Config y FileSystem.
    ' Config depende de ErrorHandler. Esto crea una dependencia circular.
    ' SOLUCIÓN: La inicialización debe ser perezosa o las dependencias deben ser ajustadas.
    ' Por ahora, para que compile, lo dejamos así, pero es un punto a revisar.
    ' Esta factoría NO debe depender de modConfigFactory.
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' La dependencia de IConfig se resolverá dentro de CErrorHandlerService
    ' para romper el ciclo.
    errorHandlerImpl.Initialize fs
    
    Set CreateErrorHandlerService = errorHandlerImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modErrorHandlerFactory.CreateErrorHandlerService: " & Err.Description
    Set CreateErrorHandlerService = Nothing
End Function


