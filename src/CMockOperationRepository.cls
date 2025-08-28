Attribute VB_Name = "CMockOperationRepository"
Option Compare Database
Option Explicit
Implements IOperationRepository

Private Sub IOperationRepository_SaveLog(ByVal operationType As String, ByVal entityId As String, ByVal details As String)
    ' Mock: No hace nada.
End Sub