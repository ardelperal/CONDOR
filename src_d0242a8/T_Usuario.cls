VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_Usuario
' DESCRIPCI├ôN: Tipo de datos para almacenar informaci├│n de usuarios del sistema
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas bas├índose en especificaci├│n funcional y servicios de autenticaci├│n
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES P├ÜBLICAS
' Basadas en la especificaci├│n funcional secci├│n 2.1 y servicios de autenticaci├│n
'******************************************************************************

' Identificador ├║nico del usuario
Public ID As Long

' Correo electr├│nico del usuario (identificador ├║nico)
Public Email As String

' Nombre completo del usuario
Public NombreCompleto As String

' Rol del usuario en el sistema CONDOR
Public Rol As E_UserRole

' Estado del usuario (Activo/Inactivo)
Public Estado As String

' Indica si es administrador (campo EsAdministrador de TbUsuariosAplicaciones)
Public EsAdministrador As Boolean

' Indica si es usuario de calidad (campo EsUsuarioCalidad de TbUsuariosAplicacionesPermisos)
Public EsUsuarioCalidad As Boolean

' Indica si es usuario t├®cnico (campo EsUsuarioTecnico de TbUsuariosAplicacionesPermisos)
Public EsUsuarioTecnico As Boolean

' ID de aplicaci├│n en sistema Lanzadera (231 para CONDOR)
Public IDAplicacion As Long

' Versi├│n instalada de la aplicaci├│n
Public VersionInstalada As String

' Fecha del ├║ltimo acceso
Public UltimoAcceso As Date

' Fecha de registro en el sistema
Public FechaRegistro As Date

' Permisos del usuario (para compatibilidad con mocks)
Public Permisos As String

'******************************************************************************
' M├ëTODOS P├ÜBLICOS
'******************************************************************************

' Valida si el usuario tiene datos m├¡nimos requeridos
Public Function EsValido() As Boolean
    EsValido = (Len(Email) > 0 And _
                InStr(Email, "@") > 1 And _
                Rol <> Rol_Desconocido And _
                Estado = "Activo")
End Function

' Determina si el usuario tiene permisos de administrador
Public Function TienePermisosAdmin() As Boolean
    TienePermisosAdmin = (EsAdministrador Or Rol = Rol_Admin)
End Function

' Determina si el usuario puede crear solicitudes
Public Function PuedeCrearSolicitudes() As Boolean
    PuedeCrearSolicitudes = (EsUsuarioCalidad Or Rol = Rol_Calidad Or TienePermisosAdmin())
End Function

' Determina si el usuario puede consultar solicitudes
Public Function PuedeConsultarSolicitudes() As Boolean
    PuedeConsultarSolicitudes = (EsUsuarioTecnico Or Rol = Rol_Tecnico Or _
                                EsUsuarioCalidad Or Rol = Rol_Calidad Or _
                                TienePermisosAdmin())
End Function
