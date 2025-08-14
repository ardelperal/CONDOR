Attribute VB_Name = "modTypes"
Option Compare Database
Option Explicit

' ============================================================================
' Módulo: modTypes
' Descripción: Definición de tipos de datos personalizados para CONDOR
' Autor: CONDOR-Expert
' Fecha: Diciembre 2024
' ============================================================================

' Tipo que representa un expediente basado en la consulta SQL de integración
' con la aplicación de Expedientes existente
Public Type T_Expediente
    IDExpediente As Long
    Nemotecnico As String
    Titulo As String
    ResponsableCalidad As String
    ResponsableTecnico As String  ' Se refiere al Jefe de Proyecto
    Pecal As String
End Type