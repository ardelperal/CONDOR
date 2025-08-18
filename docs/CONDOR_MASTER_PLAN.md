CONDOR - Plan Maestro del Proyecto
Versión: 2.3
Última Actualización: 2025-08-18

1. Visión y Objetivo Principal
El objetivo principal de CONDOR es ser la herramienta centralizada para la gestión integral del ciclo de vida de las solicitudes de Cambios, Concesiones y Desviaciones. Para ello, la aplicación debe cumplir con cuatro funcionalidades esenciales:

Registro Centralizado: Servir como el único punto de entrada para registrar los tres tipos de solicitudes: Propuestas de Cambio (PC), Concesiones/Desviaciones (CD-CA) y Concesiones/Desviaciones de Sub-suministrador (CD-CA-SUB).

Generación de Documentos (Escritura): Automatizar la generación de la documentación oficial asociada a cada tipo de solicitud, utilizando los datos registrados en el sistema para rellenar las plantillas Word predefinidas.

Sincronización de Documentos (Lectura): Permitir la actualización de los registros en la base de datos a partir de una plantilla Word que haya sido rellenada o modificada fuera de la aplicación, asegurando la consistencia de los datos.

Trazabilidad de Estado: Proporcionar una visión clara y en tiempo real del estado en el que se encuentra cada solicitud a lo largo de su ciclo de vida, desde el registro hasta el cierre.

2. Arquitectura y Principios Fundamentales
2.1. Arquitectura General
El sistema sigue una arquitectura en 3 Capas sobre un entorno Cliente-Servidor con bases de datos Access separadas para el frontend y el backend.

Capa de Presentación: Formularios de Access (.accde).

Capa de Lógica de Negocio: Clases y Módulos VBA con lógica de negocio.

Capa de Datos: Módulos VBA que gestionan el acceso a la base de datos CONDOR_datos.accdb.

2.2. Principios de Diseño (No Negociables)
Inversión de Dependencias: Las clases de alto nivel deben depender de Interfaces (I*), no de clases concretas (C*). Esto es clave para el testing y el bajo acoplamiento.

Nomenclatura Estricta:

Interfaces: IAuthService

Clases: CAuthService

Módulos: modDatabase

Tipos de Datos: T_Usuario

Miembros: camelCase (sin guiones bajos).

Testing contra la Interfaz: En los módulos de prueba (Test_*), las variables de servicio siempre se declaran del tipo de la interfaz.

3. Flujo de Trabajo y Gestión de Estados
El ciclo de vida de una solicitud se gestiona a través de un sistema de Fases y Estados controlados por el WorkflowService.

FASE REGISTRO (Estado: REGISTRADO)

FASE DESARROLLO (Estado: EN_DESARROLLO)

FASE GESTIÓN EXTERNA (Estados: REVISION_INTERNA, VALIDADO, etc.)

FASE CIERRE (Estados: APROBADA, CERRADA, RECHAZADA)

4. Especificaciones de Integración Clave
4.1. Autenticación y Roles
La autenticación se realiza a través de la Lanzadera. CONDOR recibe el email del usuario y consulta la base de datos Lanzadera_Datos.accdb para determinar el rol (Administrador, Calidad, Ingeniería) basándose en el IDAplicacion = 231.

4.2. Notificaciones Asíncronas
El sistema no envía correos directamente. En su lugar, encola las notificaciones insertando un registro en la tabla TbCorreosEnviados de la base de datos Correos_datos.accdb. Un proceso externo se encarga del envío.

Ruta Oficina: \\datoste\APLICACIONES_DYS\Aplicaciones PpD\00Recursos\Correos_datos.accdb

Ruta Local: ./back/Correos_datos.accdb

5. Estructura de la Base de Datos (CONDOR_datos.accdb)
La base de datos se compone de tablas principales para las solicitudes, tablas de workflow, tablas de logging y una tabla de mapeo para la generación de documentos.

Para un detalle exhaustivo de la estructura de las tablas, consultar el Anexo A.

Para el mapeo de campos específico para la generación de documentos, consultar el Anexo B.

6. Ciclo de Trabajo de Desarrollo (TDD Asistido)
Este es el proceso estándar para cualquier tarea de desarrollo o corrección.

Análisis y Prompt (Oráculo): El Arquitecto (CONDOR-Expert) genera un prompt detallado.

Revisión de Lecciones Aprendidas (IA): La IA debe revisar Lecciones_aprendidas.md antes de escribir código.

Desarrollo (IA): La IA implementa la funcionalidad siguiendo TDD (Tests primero).

rebuild y Pausa (IA): La IA ejecuta cscript //nologo condor_cli.vbs rebuild y se detiene.

Verificación Manual (Supervisor): El Supervisor compila el proyecto en Access.

Pruebas y Commit (IA): Tras la luz verde, la IA ejecuta los tests y, si pasan, prepara el commit.

7. Lecciones Aprendidas (Resumen)
Interfaces en VBA: La firma de los métodos debe ser idéntica.

Tests contra la Interfaz: Declarar siempre variables como Dim miServicio As IMiServicio.

Estructura de Módulos: Las declaraciones (Dim, Public, etc.) deben ir al principio del fichero.

Flujo rebuild: El comando rebuild es la fuente de verdad. La compilación manual del Supervisor es obligatoria.

Conversión Explícita: Usar siempre CLng, CStr, etc., desde Array Variant.

Tests como Especificación: Los tests y el código de acceso a datos definen las propiedades de las clases de datos (T_*).

Integración de Tests: Cada nuevo módulo de prueba (Test_*_RunAll) debe ser añadido a modTestRunner.bas.

(Este es un resumen. El documento completo Lecciones_aprendidas.md contiene más detalles).

8. Anexo A: Estructura Detallada de la Base de Datos
8.1. Tb_Solicitudes
Campo

Tipo de Dato

Descripción

ID_Solicitud

Autonumérico

Clave Primaria

ID_Expediente

Texto Corto

Clave Externa al sistema de expedientes

TipoSolicitud

Texto Corto

"CD/CA", "CD/CA-SUB", "PC"

SubTipoSolicitud

Texto Corto

"Desviación" o "Concesión"

CodigoSolicitud

Texto Corto

Código autogenerado para la solicitud

EstadoInterno

Texto Corto

Estado actual del workflow

FechaCreacion

Fecha/Hora

Timestamp de creación

UsuarioCreacion

Texto Corto

Email del usuario creador

8.2. Tb_Datos_PC (Propuesta de Cambio)
(Contiene todos los campos específicos del formulario F4203.11)

8.3. Tb_Datos_CD_CA (Desviación / Concesión)
(Contiene todos los campos específicos del formulario F4203.10)

8.4. Tb_Datos_CD_CA_SUB (Desviación / Concesión Sub-suministrador)
(Contiene todos los campos específicos del formulario F4203.101)

8.5. Tb_Mapeo_Campos
Tabla pre-poblada que contiene el mapeo entre los campos de las tablas de datos y los marcadores en las plantillas Word.
| Campo | Tipo de Dato |
| :--- | :--- |
| ID_Mapeo | Autonumérico |
| NombrePlantilla| Texto Corto |
| NombreCampoTabla| Texto Corto |
| ValorAsociado| Texto Corto |
| NombreCampoWord| Texto Corto |

8.6. Tablas de Soporte
Tb_Log_Cambios: Auditoría de cambios en el sistema.

Tb_Log_Errores: Registro de errores de la aplicación.

Tb_Adjuntos: Gestión de ficheros adjuntos a las solicitudes.

TbEstados: Definición de los estados del workflow.

TbTransiciones: Reglas para las transiciones de estado permitidas.

9. Anexo B: Mapeo de Campos para Generación de Documentos
9.1. Plantilla "PC" (F4203.11 - Propuesta de Cambio)
NombrePlantilla

NombreCampoTabla (en Tb_Datos_PC)

ValorAsociado

NombreCampoWord

"PC"

RefContratoInspeccionOficial

NULL

Parte0_1

"PC"

RefSuministrador

NULL

Parte0_2

"PC"

SuministradorNombreDir

NULL

Parte1_1

"PC"

ObjetoContrato

NULL

Parte1_2

"PC"

DescripcionMaterialAfectado

NULL

Parte1_3

"PC"

NumPlanoEspecificacion

NULL

Parte1_4

"PC"

DescripcionPropuestaCambio

NULL

Parte1_5

"PC"

DescripcionPropuestaCambio_Cont

NULL

Parte1_5Cont

"PC"

Motivo_CorregirDeficiencias

True

Parte1_6_1

"PC"

Motivo_MejorarCapacidad

True

Parte1_6_2

"PC"

Motivo_AumentarNacionalizacion

True

Parte1_6_3

"PC"

Motivo_MejorarSeguridad

True

Parte1_6_4

"PC"

Motivo_MejorarFiabilidad

True

Parte1_6_5

"PC"

Motivo_MejorarCosteEficacia

True

Parte1_6_6

"PC"

Motivo_Otros

True

Parte1_6_7

"PC"

Motivo_Otros_Detalle

NULL

Parte1_6_8

"PC"

IncidenciaCoste

"Aumentará"

Parte1_7a_1

"PC"

IncidenciaCoste

"Disminuirá"

Parte1_7a_2

"PC"

IncidenciaCoste

"No variará"

Parte1_7a_3

"PC"

IncidenciaPlazo

"Aumentará"

Parte1_7b_1

"PC"

IncidenciaPlazo

"Disminuirá"

Parte1_7b_2

"PC"

IncidenciaPlazo

"No variará"

Parte1_7b_3

"PC"

Incidencia_Seguridad

True

Parte1_7c_1

"PC"

Incidencia_Fiabilidad

True

Parte1_7c_2

"PC"

Incidencia_Mantenibilidad

True

Parte1_7c_3

"PC"

Incidencia_Intercambiabilidad

True

Parte1_7c_4

"PC"

Incidencia_VidaUtilAlmacen

True

Parte1_7c_5

"PC"

Incidencia_FuncionamientoFuncion

True

Parte1_7c_6

"PC"

CambioAfecta_MaterialEntregado

True

Parte1_9_1

"PC"

CambioAfecta_MaterialPorEntregar

True

Parte1_9_2

"PC"

FirmaOficinaTecnica_Nombre

NULL

Parte1_10

"PC"

FirmaRepSuministrador_Nombre

NULL

Parte1_11

"PC"

ObservacionesRAC_Ref

NULL

Parte2_1

"PC"

RAC_Codigo

NULL

Parte2_2

"PC"

ObservacionesRAC

NULL

Parte2_3

"PC"

FechaFirmaRAC

NULL

Parte2_4

"PC"

ObsAprobacionAutoridadDiseno

NULL

Parte3_1

"PC"

FirmaAutoridadDiseno_NombreCargo

NULL

Parte3_2

"PC"

FechaFirmaAutoridadDiseno

NULL

Parte3_3

"PC"

DecisionFinal

"APROBADO"

Parte3_2_1

"PC"

DecisionFinal

"NO APROBADO"

Parte3_2_2

"PC"

ObsDecisionFinal

NULL

Parte3_3_1

"PC"

CargoFirmanteFinal

NULL

Parte3_3_2

"PC"

FechaFirmaDecisionFinal

NULL

Parte3_3_3

9.2. Plantilla "CDCA" (F4203.10 - Desviación / Concesión)
NombrePlantilla

NombreCampoTabla (en Tb_Datos_CD_CA)

ValorAsociado

NombreCampoWord

"CDCA"

RefSuministrador

NULL

Parte0_1

"CDCA"

NumContrato

NULL

Parte1_2

"CDCA"

IdentificacionMaterial

NULL

Parte1_3

"CDCA"

NumPlanoEspecificacion

NULL

Parte1_4

"CDCA"

CantidadPeriodo

NULL

Parte1_5a

"CDCA"

NumSerieLote

NULL

Parte1_5b

"CDCA"

DescripcionImpactoNC

NULL

Parte1_6

"CDCA"

RefDesviacionesPrevias

NULL

Parte1_7

"CDCA"

CausaNC

NULL

Parte1_8

"CDCA"

ImpactoCoste

"Increased / aumentado"

Parte1_9_1

"CDCA"

ImpactoCoste

"Decreased / disminuido"

Parte1_9_2

"CDCA"

ImpactoCoste

"Unchanged / sin cambio"

Parte1_9_3

"CDCA"

ClasificacionNC

"Major / Mayor"

Parte1_10_1

"CDCA"

ClasificacionNC

"Minor / Menor"

Parte1_10_2

"CDCA"

RequiereModificacionContrato

True

Parte1_12_1

"CDCA"

EfectoFechaEntrega

NULL

Parte1_13

"CDCA"

IdentificacionAutoridadDiseno

NULL

Parte1_14

"CDCA"

EsSuministradorAD

True

Parte1_18_1

"CDCA"

EsSuministradorAD

False

Parte1_18_2

"CDCA"

DescripcionImpactoNC_Cont

NULL

Parte1_20

"CDCA"

RAC_Ref

NULL

Parte2_21_1

"CDCA"

RAC_Codigo

NULL

Parte2_21_2

"CDCA"

ObservacionesRAC

NULL

Parte2_21_3

"CDCA"

FechaFirmaRAC

NULL

Parte2_22

"CDCA"

DecisionFinal

"APROBADO"

Parte3_23_1

"CDCA"

DecisionFinal

"NO APROBADO"

Parte3_23_2

"CDCA"

ObservacionesFinales

NULL

Parte3_24_1

"CDCA"

FechaFirmaDecisionFinal

NULL

Parte3_24_2

"CDCA"

CargoFirmanteFinal

NULL

Parte3_24_4

9.3. Plantilla "CDCASUB" (F4203.101 - Desviación / Concesión Sub-suministrador)
NombrePlantilla

NombreCampoTabla (en Tb_Datos_CD_CA_SUB)

ValorAsociado

NombreCampoWord

"CDCASUB"

RefSuministrador

NULL

Parte0_1

"CDCASUB"

RefSubSuministrador

NULL

Parte0_2

"CDCASUB"

SuministradorPrincipalNombreDir

NULL

Parte1_1

"CDCASUB"

SubSuministradorNombreDir

NULL

Parte1_2

"CDCASUB"

IdentificacionMaterial

NULL

Parte1_5

"CDCASUB"

NumPlanoEspecificacion

NULL

Parte1_6

"CDCASUB"

CantidadPeriodo

NULL

Parte1_7a

"CDCASUB"

NumSerieLote

NULL

Parte1_7b

"CDCASUB"

DescripcionImpactoNC

NULL

Parte1_8

"CDCASUB"

RefDesviacionesPrevias

NULL

Parte1_9

"CDCASUB"

CausaNC

NULL

Parte1_10

"CDCASUB"

ImpactoCoste

"Incrementado"

Parte1_11_1

"CDCASUB"

ImpactoCoste

"Sin cambio"

Parte1_11_2

"CDCASUB"

ImpactoCoste

"Disminuido"

Parte1_11_3

"CDCASUB"

ClasificacionNC

"Mayor"

Parte1_12_1

"CDCASUB"

ClasificacionNC

"Menor"

Parte1_12_2

"CDCASUB"

Afecta_Prestaciones

True

Parte1_13_1

"CDCASUB"

Afecta_Seguridad

True

Parte1_13_2

"CDCASUB"

Afecta_Fiabilidad

True

Parte1_13_3

"CDCASUB"

Afecta_VidaUtil

True

Parte1_13_4

"CDCASUB"

Afecta_Medioambiente

True

Parte1_13_5

"CDCASUB"

Afecta_Intercambiabilidad

True

Parte1_13_6

"CDCASUB"

Afecta_Mantenibilidad

True

Parte1_13_7

"CDCASUB"

Afecta_Apariencia

True

Parte1_13_8

"CDCASUB"

Afecta_Otros

True

Parte1_13_9

"CDCASUB"

RequiereModificacionContrato

True

Parte1_14

"CDCASUB"

EfectoFechaEntrega

NULL

Parte1_15

"CDCASUB"

IdentificacionAutoridadDiseno

NULL

Parte1_16

"CDCASUB"

EsSubSuministradorAD

True

Parte1_20_1

"CDCASUB"

EsSubSuministradorAD

False

Parte1_20_2

"CDCASUB"

NombreRepSubSuministrador

NULL

Parte1_21

"CDCASUB"

DescripcionImpactoNC_Cont

NULL

Parte1_22

"CDCASUB"

RAC_Ref

NULL

Parte2_23_1

"CDCASUB"

RAC_Codigo

NULL

Parte2_23_2

"CDCASUB"

ObservacionesRAC

NULL

Parte2_23_3

"CDCASUB"

FechaFirmaRAC

NULL

Parte2_25

"CDCASUB"

DecisionSuministradorPrincipal

"APROBADO"

Parte3_26_1

"CDCASUB"

DecisionSuministradorPrincipal

"NO APROBADO"

Parte3_26_2

"CDCASUB"

ObsSuministradorPrincipal

NULL

Parte3_27_1

"CDCASUB"

FechaFirmaSuministradorPrincipal

NULL

Parte3_27_2

"CDCASUB"

FirmaSuministradorPrincipal_NombreCargo

NULL

Parte3_27_4

"CDCASUB"

ObsRACDelegador

NULL

Parte4_28

"CDCASUB"

FechaFirmaRACDelegador

NULL

Parte4_30

