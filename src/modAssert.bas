Attribute VB_Name = "modAssert"
' Modulo: modAssert
' Proposito: Funciones de asercion para las pruebas.
Option Compare Database
Option Explicit

Public Sub IsTrue(value As Boolean, message As String)
    If Not value Then
        Debug.Print "ASSERT FAILED: " & message
    End If
End Sub