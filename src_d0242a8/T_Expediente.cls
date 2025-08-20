VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' CLASE: T_Expediente
' DESCRIPCI├ôN: Tipo de datos para almacenar informaci├│n de expedientes
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas bas├índose en CExpedienteService.cls y CMockExpedienteService.cls como fuente de verdad
'******************************************************************************

Option Compare Database
Option Explicit

'******************************************************************************
' PROPIEDADES P├ÜBLICAS
' Basadas en el an├ílisis de CExpedienteService.cls y CMockExpedienteService.cls
'******************************************************************************

' Identificador ├║nico del expediente
Public IDExpediente As Long

' Nemot├®cnico del expediente
Public Nemotecnico As String

' T├¡tulo del expediente
Public Titulo As String

' Responsable de calidad del expediente
Public ResponsableCalidad As String

' Responsable t├®cnico del expediente
Public ResponsableTecnico As String

' C├│digo PECAL del expediente
Public Pecal As String

'******************************************************************************
' PROPIEDADES ADICIONALES IDENTIFICADAS EN DOCUMENTACI├ôN
'******************************************************************************

' N├║mero de expediente (string)
Public NumeroExpediente As String

' Estado del expediente
Public EstadoExpediente As String

' Email del responsable
Public EmailResponsable As String

' Jefe de proyecto
Public JefeProyecto As String

' Contratista principal
Public ContratistaPrincipal As String

' Valor del contrato
Public ValorContrato As Currency

' Fecha de inicio
Public FechaInicio As Date

' Fecha de fin prevista
Public FechaFinPrevista As Date

' ├Ültima actualizaci├│n
Public UltimaActualizacion As Date

'******************************************************************************
' PROPIEDADES CR├ìTICAS IDENTIFICADAS EN TESTS
'******************************************************************************

' Identificador ├║nico usado en modMockFramework
Public ID As Long

' Fecha de creaci├│n usado en Test_CExpedienteService.bas
Public FechaCreacion As Date

' Estado usado en Test_CExpedienteService.bas
Public Estado As String

'******************************************************************************
' M├ëTODOS P├ÜBLICOS
'******************************************************************************

' M├®todo para validar completitud de datos
Public Function EsValido() As Boolean
    EsValido = (Len(NumeroExpediente) > 0 And _
                Len(ResponsableCalidad) > 0 And _
                Len(EmailResponsable) > 0)
End Function
