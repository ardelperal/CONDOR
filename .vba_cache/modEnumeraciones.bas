Attribute VB_Name = "modEnumeraciones"
Option Compare Database
Option Explicit


' Enumeración de roles de usuario para el sistema CONDOR
' Define los diferentes tipos de roles disponibles
Public Enum EUserRole
    Rol_Desconocido = 0
    ROL_ADMINISTRADOR = 1
    Rol_Calidad = 2
    Rol_Tecnico = 3
End Enum

