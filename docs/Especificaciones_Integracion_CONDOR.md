Especificaciones Técnicas de Integración - Sistema CONDOR
Introducción
Este documento detalla las especificaciones técnicas para la integración del Sistema CONDOR con los sistemas externos, basado en la arquitectura definida en la especificación funcional del proyecto.

1. Integración con ExpedienteService
1.1 Descripción General
El ExpedienteService es el servicio principal para obtener información de expedientes desde el sistema externo de gestión de contratos. CONDOR utiliza este servicio para:

Obtener datos básicos del expediente al crear solicitudes

Validar la existencia y estado de expedientes

Sincronizar información de responsables y contratistas

1.2 Arquitectura de Integración
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────────┐
│   CONDOR.accdb  │    │  ExpedienteService │    │ Sistema Expedientes │
│   (Frontend)    │◄──►│   (Capa Servicio) │◄──►│   (Base de Datos)   │
└─────────────────┘    └──────────────────┘    └─────────────────────┘

1.3 Interfaz IExpedienteService
Definición de la Interfaz
' Archivo: IExpedienteService.cls
' Descripción: Interfaz para el servicio de expedientes

Option Compare Database
Option Explicit

' Método principal para obtener datos de expediente
Public Function ObtenerExpediente(numeroExpediente As String) As T_Expediente
End Function

' Método para validar existencia de expediente
Public Function ExisteExpediente(numeroExpediente As String) As Boolean
End Function

' Método para obtener lista de expedientes por responsable
Public Function ObtenerExpedientesPorResponsable(emailResponsable As String) As Collection
End Function

2. Integración con Servicio de Notificaciones Asíncrono
2.1 Descripción General
El sistema de notificaciones no envía correos electrónicos directamente. En su lugar, opera como una cola de trabajo mediante la inserción de registros en una base de datos compartida.

Responsabilidad de CONDOR: La única responsabilidad del NotificationService de CONDOR es crear y guardar un registro en la tabla TbCorreosEnviados de la base de datos Correos_datos.accdb.

Proceso Asíncrono: Un proceso externo y centralizado es responsable de leer los registros pendientes (donde FechaEnvio es Nulo), enviar los correos vía SMTP y actualizar el registro con la FechaEnvio.

2.2 Arquitectura de Integración
┌─────────────────┐      ┌─────────────────────┐      ┌───────────────────┐
│  CONDOR.accdb   │───►│ Correos_datos.accdb │◄───│  Proceso Cíclico  │
│ (Escribe en la cola) │ (Tabla TbCorreosEnviados) │ (Lee y envía correos) │
└─────────────────┘      └─────────────────────┘      └───────────────────┘

2.3 Ubicación de la Base de Datos de Correos
La ruta a la base de datos Correos_datos.accdb depende del entorno de ejecución y debe ser gestionada por el ConfigService.

Entorno Oficina: \\datoste\APLICACIONES_DYS\Aplicaciones PpD\00Recursos\Correos_datos.accdb

Entorno Local: ./back/Correos_datos.accdb

2.4 Estructura de la Tabla TbCorreosEnviados
Campo

Regla de Negocio para CONDOR

IDCorreo

Se debe calcular como DMax("IDCorreo", "TbCorreosEnviados") + 1.

Aplicacion

Siempre debe ser el valor fijo "CONDOR".

Asunto

El asunto del correo.

Cuerpo

El cuerpo del correo, formateado en HTML atractivo.

Destinatarios

Lista de correos de los destinatarios principales, separados por punto y coma.

DestinatariosConCopia

Siempre debe ser el correo del usuario actual en sesión.

DestinatariosConCopiaOculta

Siempre debe ser el correo del administrador del sistema (obtenido de la configuración).

URLAdjunto

Ruta completa a un fichero adjunto, si lo hubiera.

FechaGrabacion

El timestamp del momento exacto de la inserción del registro.

FechaEnvio

Siempre se debe dejar Nulo.

2.5 Interfaz del Servicio (INotificationService)
' INotificationService.cls
Public Function EnviarNotificacion( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
End Function

2.6 Implementación del Servicio (CNotificationService)
La implementación debe:

1. Obtener la ruta de Correos_datos.accdb desde ConfigService
2. Calcular automáticamente IDCorreo como DMax("IDCorreo", "TbCorreosEnviados") + 1
3. Establecer Aplicacion = "CONDOR"
4. Obtener el correo del usuario actual para DestinatariosConCopia
5. Obtener el correo del administrador para DestinatariosConCopiaOculta
6. Establecer FechaGrabacion = Now()
7. Dejar FechaEnvio como Null
8. Insertar el registro en TbCorreosEnviados usando la conexión externa

3. Integración con Sistema RAC
3.1 Descripción General
El Sistema RAC (Registro y Control) es el sistema externo donde se registran oficialmente las solicitudes aprobadas. La integración permite:

Envío automático de solicitudes aprobadas

Sincronización de estados entre CONDOR y RAC

Obtención de números de registro oficiales

3.2 Flujo de Integración
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│ CONDOR          │    │ Servicio RAC    │    │ Sistema RAC     │
│ Estado: Aprobado│───►│ Envío Automático│───►│ Registro Oficial│
└─────────────────┘    └─────────────────┘    └─────────────────┘
         ▲                       │                       │
         │                       ▼                       │
         └─────────────── Confirmación ◄─────────────────┘
                        + Número RAC

4. Sistema de Lanzadera
4.1 Descripción General
El Sistema de Lanzadera gestiona el despliegue y actualización automática de CONDOR. Componentes principales:

Lanzadera_Datos.accdb: Base de datos de control de versiones y usuarios

condor_cli.vbs: Herramienta de línea de comandos para operaciones

Sistema de actualización automática

5. Protocolos de Comunicación
5.1 Formato de Mensajes
Todos los servicios utilizan formato JSON para intercambio de datos:

{
  "version": "1.0",
  "timestamp": "2024-12-20T10:30:00Z",
  "source": "CONDOR",
  "target": "ExpedienteService",
  "operation": "ObtenerExpediente",
  "data": {
    "numeroExpediente": "EXP-2024-INF-001"
  },
  "metadata": {
    "usuario": "maria.garcia@empresa.com",
    "sessionId": "sess_123456789"
  }
}
