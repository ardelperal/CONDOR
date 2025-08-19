VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Datos_PC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Option Explicit

' Clase de tipo de datos para tbDatosPC
' Basada en el Anexo A del CONDOR_MASTER_PLAN.md

Private m_idDatosPC As Long
Private m_refContratoInspeccionOficial As String
Private m_refContratoInspeccionOficialNuevo As String
Private m_refSuministrador As String
Private m_refSuministradorNuevo As String
Private m_refSubSuministrador As String
Private m_refSubSuministradorNuevo As String
Private m_importeOriginal As Currency
Private m_importeNuevo As Currency
Private m_justificacion As String
Private m_descripcionCambio As String
Private m_impactoCalidad As String
Private m_impactoCoste As String
Private m_impactoPlazos As String
Private m_requiereModificacionContrato As Boolean
Private m_efectoFechaEntrega As String
Private m_observacionesAdicionales As String

Public Property Get idDatosPC() As Long
    idDatosPC = m_idDatosPC
End Property

Public Property Let idDatosPC(ByVal value As Long)
    m_idDatosPC = value
End Property

Public Property Get refContratoInspeccionOficial() As String
    refContratoInspeccionOficial = m_refContratoInspeccionOficial
End Property

Public Property Let refContratoInspeccionOficial(ByVal value As String)
    m_refContratoInspeccionOficial = value
End Property

Public Property Get refContratoInspeccionOficialNuevo() As String
    refContratoInspeccionOficialNuevo = m_refContratoInspeccionOficialNuevo
End Property

Public Property Let refContratoInspeccionOficialNuevo(ByVal value As String)
    m_refContratoInspeccionOficialNuevo = value
End Property

Public Property Get refSuministrador() As String
    refSuministrador = m_refSuministrador
End Property

Public Property Let refSuministrador(ByVal value As String)
    m_refSuministrador = value
End Property

Public Property Get refSuministradorNuevo() As String
    refSuministradorNuevo = m_refSuministradorNuevo
End Property

Public Property Let refSuministradorNuevo(ByVal value As String)
    m_refSuministradorNuevo = value
End Property

Public Property Get refSubSuministrador() As String
    refSubSuministrador = m_refSubSuministrador
End Property

Public Property Let refSubSuministrador(ByVal value As String)
    m_refSubSuministrador = value
End Property

Public Property Get refSubSuministradorNuevo() As String
    refSubSuministradorNuevo = m_refSubSuministradorNuevo
End Property

Public Property Let refSubSuministradorNuevo(ByVal value As String)
    m_refSubSuministradorNuevo = value
End Property

Public Property Get importeOriginal() As Currency
    importeOriginal = m_importeOriginal
End Property

Public Property Let importeOriginal(ByVal value As Currency)
    m_importeOriginal = value
End Property

Public Property Get importeNuevo() As Currency
    importeNuevo = m_importeNuevo
End Property

Public Property Let importeNuevo(ByVal value As Currency)
    m_importeNuevo = value
End Property

Public Property Get justificacion() As String
    justificacion = m_justificacion
End Property

Public Property Let justificacion(ByVal value As String)
    m_justificacion = value
End Property

Public Property Get descripcionCambio() As String
    descripcionCambio = m_descripcionCambio
End Property

Public Property Let descripcionCambio(ByVal value As String)
    m_descripcionCambio = value
End Property

Public Property Get impactoCalidad() As String
    impactoCalidad = m_impactoCalidad
End Property

Public Property Let impactoCalidad(ByVal value As String)
    m_impactoCalidad = value
End Property

Public Property Get impactoCoste() As String
    impactoCoste = m_impactoCoste
End Property

Public Property Let impactoCoste(ByVal value As String)
    m_impactoCoste = value
End Property

Public Property Get impactoPlazos() As String
    impactoPlazos = m_impactoPlazos
End Property

Public Property Let impactoPlazos(ByVal value As String)
    m_impactoPlazos = value
End Property

Public Property Get requiereModificacionContrato() As Boolean
    requiereModificacionContrato = m_requiereModificacionContrato
End Property

Public Property Let requiereModificacionContrato(ByVal value As Boolean)
    m_requiereModificacionContrato = value
End Property

Public Property Get efectoFechaEntrega() As String
    efectoFechaEntrega = m_efectoFechaEntrega
End Property

Public Property Let efectoFechaEntrega(ByVal value As String)
    m_efectoFechaEntrega = value
End Property

Public Property Get observacionesAdicionales() As String
    observacionesAdicionales = m_observacionesAdicionales
End Property

Public Property Let observacionesAdicionales(ByVal value As String)
    m_observacionesAdicionales = value
End Property





