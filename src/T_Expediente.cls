VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

'******************************************************************************
' CLASE: T_Expediente
' DESCRIPCIÓN: Tipo de datos para almacenar información de expedientes
' AUTOR: Sistema CONDOR
' FECHA: 2024
' NOTA: Propiedades definidas basándose en CExpedienteService.cls y CMockExpedienteService.cls como fuente de verdad
'******************************************************************************


'******************************************************************************
' PROPIEDADES PÚBLICAS
' Basadas en el análisis de CExpedienteService.cls y CMockExpedienteService.cls
'******************************************************************************

' Identificador único del expediente
Public idExpediente As Long

' Nemotécnico del expediente
Public Nemotecnico As String

' Título del expediente
Public Titulo As String

' Responsable de calidad del expediente
Public ResponsableCalidad As String

' Responsable técnico del expediente
Public ResponsableTecnico As String

' Código PECAL del expediente
Public Pecal As String

'******************************************************************************
' PROPIEDADES ADICIONALES IDENTIFICADAS EN DOCUMENTACIÓN
'******************************************************************************

' Número de expediente (string)
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

' Última actualización
Public UltimaActualizacion As Date

'******************************************************************************
' PROPIEDADES CRÍTICAS IDENTIFICADAS EN TESTS
'******************************************************************************

' Identificador único usado en modMockFramework
Public ID As Long

' Fecha de creación usado en Test_CExpedienteService.bas
Public fechaCreacion As Date

' Estado usado en Test_CExpedienteService.bas
Public Estado As String

'******************************************************************************
' MÉTODOS PÚBLICOS
'******************************************************************************

' Método para validar completitud de datos
Public Function EsValido() As Boolean
    EsValido = (Len(NumeroExpediente) > 0 And _
                Len(ResponsableCalidad) > 0 And _
                Len(EmailResponsable) > 0)
End Function











