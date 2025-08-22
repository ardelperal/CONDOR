VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Clase: CTestResult
' Proposito: Almacena el resultado de un unico test.
Option Compare Database
Option Explicit

Public TestName As String
Public Success As Boolean
Public ErrorMessage As String