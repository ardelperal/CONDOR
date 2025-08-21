
Visión y Objetivo Principal

1. El objetivo principal de CONDOR es ser la herramienta centralizada para la gestión integral del ciclo de vida de las solicitudes de Cambios, Concesiones y Desviaciones. Para ello, la aplicación debe cumplir con cuatro funcionalidades esenciales:

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
   El flujo de trabajo de la aplicación se divide en fases gestionadas por los roles Calidad y Técnico. El rol Administrador tiene acceso a todas las funcionalidades.

Fase 1: Registro (A cargo de Calidad)
Inicio: Un usuario con rol Calidad inicia el proceso de "Alta de Solicitud".

Selección de Expediente: El usuario elige un expediente de una lista precargada desde la base de datos de Expedientes.

Selección de Suministrador: Se selecciona un suministrador asociado al expediente elegido.

Selección de Tipo de Solicitud: Calidad elige si la solicitud es de tipo PC o CD-CA.

Lógica de Sub-contratista: Si se elige CD-CA, el sistema consulta el campo ContratistaPrincipal del expediente. Si el valor es 'Sí', la solicitud se clasifica como CD-CA; en caso contrario, se clasifica como CD-CA-SUB.

Cumplimentación Inicial: Calidad rellena los campos iniciales de la solicitud.

Pase a Técnico: Al guardar, la solicitud entra en la FASE DE REGISTRO. El sistema automáticamente:

Rellena el campo fechaPaseTecnico en la tabla tbSolicitudes.

Encola una notificación por correo electrónico para el equipo Técnico responsable de ese expediente.

Fase 2: Desarrollo Técnico (A cargo del Técnico)
Recepción: Un usuario con rol Técnico accede a su "bandeja de entrada", que muestra las solicitudes asociadas a sus expedientes y que están en la fase técnica (es decir, tienen fechaPaseTecnico pero no fechaCompletadoTecnico).

Cumplimentación Técnica: El técnico rellena los campos técnicos correspondientes a la solicitud.

Liberación: Una vez completada su parte, el técnico pulsa un botón de "Liberar" o "Finalizar". El sistema automáticamente:

Rellena el campo fechaCompletadoTecnico en la tabla tbSolicitudes.

Encola una notificación por correo electrónico para el usuario de Calidad que inició el proceso.

Fase 3: Gestión Externa y Cierre (A cargo de Calidad)
Recepción: El usuario de Calidad recibe la notificación y ve en su panel que la solicitud ha vuelto de la fase técnica.

Generación de Documentos: Calidad utiliza CONDOR para generar la plantilla Word (.docx) con los datos de la solicitud. Cada versión del documento generado se guarda en un directorio de anexos para mantener la trazabilidad.

Interacción Externa (Fuera de CONDOR): Calidad gestiona la comunicación con los agentes externos (suministradores, etc.) por correo electrónico, enviando y recibiendo las plantillas Word.

Actualización de Datos (Sincronización): A medida que recibe las plantillas actualizadas de agentes externos, Calidad utiliza una funcionalidad específica en la interfaz de CONDOR (p. ej., un botón "Sincronizar desde Documento"). Al activarla, la aplicación:
            1.  Abre un selector de archivos para que el usuario elija el documento `.docx` actualizado.
            2.  Lee el contenido del documento Word, extrae los datos de los campos relevantes (según el mapeo del Anexo B).
            3.  Actualiza automáticamente los campos correspondientes en la base de datos de CONDOR.
            Este proceso evita la entrada manual de datos, reduce errores y asegura la consistencia.

Cierre: El proceso continúa hasta que la solicitud es finalmente aprobada o denegada, momento en el cual Calidad actualiza el estado final en el sistema.

4. Especificaciones de Integración Clave
   4.1. Autenticación y Roles
   El sistema de autenticación y autorización está centralizado y se integra con la aplicación "Lanzadera" de la oficina.

4.1.1. Flujo de Arranque
El usuario abre CONDOR desde la Lanzadera.

La Lanzadera pasa el correo electrónico del usuario logueado a CONDOR a través del parámetro VBA.Command.

4.1.2. Lógica de Determinación de Rol
CONDOR utiliza el correo electrónico recibido para determinar el rol del usuario mediante consultas a la base de datos de la Lanzadera.

Base de Datos de Roles: Lanzadera_Datos.accdb

Ruta Producción: \\datoste\aplicaciones_dys\Aplicaciones PpD\Lanzadera\Lanzadera_Datos.accdb

Ruta Local: ./back/Lanzadera_Datos.accdb

ID de Aplicación para CONDOR: 231

4.1.3. Consulta de Rol de Administrador Global
Se verifica si el usuario es un administrador global en la tabla TbUsuariosAplicaciones. Si el campo EsAdministrador es 'Sí', se asigna el rol de Administrador y el proceso finaliza.

4.1.4. Consulta de Roles Específicos de la Aplicación
Si no es administrador global, se consulta la tabla TbUsuariosAplicacionesPermisos con el email del usuario y IDAplicacion = 231 para determinar el rol (Administrador, Calidad o Técnico).

4.1.5. Seguridad de la Base de Datos
Regla Crítica: Todas las bases de datos del backend (Lanzadera_Datos.accdb, CONDOR_datos.accdb, Correos_datos.accdb, etc.), tanto en entorno de producción como local, están protegidas por contraseña.

Contraseña Universal: dpddpd

4.2. Integración con Sistema de Expedientes
4.2.1. Flujo de Trabajo y Propósito
Toda solicitud en CONDOR (PC, CD/CA, CD/CA-SUB) debe estar asociada a un Expediente. El primer paso para un usuario de Calidad al crear una nueva solicitud es seleccionar el expediente sobre el cual se va a actuar. CONDOR se conecta a una base de datos externa para listar los expedientes disponibles.

4.2.2. Base de Datos de Expedientes
Nombre: Expedientes_datos.accdb

Ruta Producción: \\datoste\aplicaciones_dys\Aplicaciones PpD\Expedientes\Expedientes_datos.accdb

Ruta Local: ./back/Expedientes_datos.accdb

4.2.3. Consultas de Selección de Expedientes
Consulta General (Rol Calidad):
Para poblar el selector de expedientes, se utiliza la siguiente consulta para mostrar solo los expedientes activos, adjudicados y que cumplen con la normativa de calidad PECAL.

SELECT
    E.IDExpediente,
    E.Nemotecnico,
    E.Titulo,
    E.CodExp,
    E.FechaInicioContrato,
    E.FechaFinContrato,
    E.FechaFinGarantia,
    U.Nombre AS ResponsableCalidad,
    E.ContratistaPrincipal
FROM
    TbExpedientes AS E LEFT JOIN TbUsuariosAplicaciones AS U
    ON E.IDResponsableCalidad = U.Id
WHERE
    E.Adjudicado='Sí' AND E.Pecal='Sí'
ORDER BY
    E.IDExpediente DESC, E.ContratistaPrincipal DESC;

Consulta por Responsable (Rol Técnico):
Para filtrar y mostrar a los usuarios técnicos solo las solicitudes de los expedientes en los que son Jefes de Proyecto o responsables.

SELECT
    E.IDExpediente,
    E.Nemotecnico,
    E.Titulo,
    E.CodExp,
    E.FechaInicioContrato,
    E.FechaFinContrato,
    E.FechaFinGarantia,
    E.ContratistaPrincipal,
    ER.EsJefeProyecto,
    U.Nombre AS JP
FROM
    (TbExpedientes AS E INNER JOIN TbExpedientesResponsables AS ER
    ON E.IDExpediente = ER.IdExpediente)
    INNER JOIN TbUsuariosAplicaciones AS U
    ON ER.IdUsuario = U.Id
WHERE
    E.Adjudicado='Sí' AND E.Pecal='Sí' AND ER.EsJefeProyecto='Sí'
ORDER BY
    E.IDExpediente DESC, E.ContratistaPrincipal DESC;

**Definición de Términos Clave:**
*   **PECAL (Publicaciones Españolas de Calidad):** Se refiere a un conjunto de normas que establecen los requisitos de aseguramiento de la calidad para empresas que suministran bienes y servicios al Ministerio de Defensa español. Estas normas son la adaptación nacional de las normas AQAP (Allied Quality Assurance Publications) de la OTAN. La condición `Pecal='Sí'` en una consulta asegura que solo se procesan expedientes que cumplen con estos estándares de calidad.

4.2.4. Alcance de la Integración
La interacción de CONDOR con la base de datos de expedientes es de solo lectura. Las únicas operaciones permitidas son:

Listar expedientes para su selección.

Tomar el IDExpediente seleccionado para usarlo como clave externa en la tabla tbSolicitudes de CONDOR.
No se crearán, modificarán ni eliminarán expedientes desde CONDOR.

4.3. Notificaciones Asíncronas
El sistema no envía correos directamente. En su lugar, encola las notificaciones insertando un registro en la tabla TbCorreosEnviados de la base de datos Correos_datos.accdb. Un proceso externo se encarga del envío.

Ruta Oficina: \\datoste\APLICACIONES_DYS\Aplicaciones PpD\00Recursos\Correos_datos.accdb

Ruta Local: ./back/Correos_datos.accdb

5. Estructura de la Base de Datos (CONDOR_datos.accdb)
   La base de datos se compone de tablas principales para las solicitudes, tablas de workflow, tablas de logging y una tabla de mapeo para la generación de documentos.

Para un detalle exhaustivo de la estructura de las tablas, consultar el Anexo A.

Para el mapeo de campos específico para la generación de documentos, consultar el Anexo B.

6. Ciclo de Trabajo de Desarrollo (TDD Asistido con Sincronización Discrecional)
   Este es el proceso estándar para cualquier tarea de desarrollo o corrección, optimizado para permitir actualizaciones selectivas de módulos.

Análisis y Prompt (Oráculo): El Arquitecto (CONDOR-Expert) genera un prompt detallado.

Revisión de Lecciones Aprendidas (IA): La IA debe revisar Lecciones_aprendidas.md antes de escribir código.

Desarrollo (IA): La IA implementa la funcionalidad siguiendo TDD (Tests primero).

Sincronización Selectiva y Pausa (IA): La IA ejecuta:
   - `cscript //nologo condor_cli.vbs update [módulos_específicos]` para cambios puntuales
   - `cscript //nologo condor_cli.vbs update` para sincronización automática optimizada (solo abre BD si hay cambios)
   - `cscript //nologo condor_cli.vbs rebuild` solo si hay problemas graves de sincronización
   
   **Nota:** Todos los comandos incluyen conversión automática UTF-8 a ANSI para soporte completo de caracteres especiales.
   Luego se detiene y espera confirmación.

Verificación Manual (Supervisor): El Supervisor compila el proyecto en Access.

Pruebas y Commit (IA): Tras la luz verde, la IA ejecuta los tests y, si pasan, prepara el commit.

**Ventajas de la Sincronización Discrecional:**
- **Eficiencia**: Solo actualiza los módulos modificados, reduciendo el tiempo de sincronización
- **Estabilidad**: Minimiza el riesgo de afectar módulos no relacionados con los cambios
- **Desarrollo Iterativo**: Facilita ciclos rápidos de desarrollo-prueba-corrección
- **Flexibilidad**: Permite trabajar en funcionalidades específicas sin impactar el proyecto completo

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
   8.1. tbSolicitudes
   Campo

Tipo de Dato

Descripción

idSolicitud

Autonumérico

Clave Primaria

idExpediente

Texto Corto

Clave Externa al sistema de expedientes

tipoSolicitud

Texto Corto

"CD/CA", "CD/CA-SUB", "PC"

subTipoSolicitud

Texto Corto

"Desviación" o "Concesión"

codigoSolicitud

Texto Corto

Código autogenerado para la solicitud

estadoInterno

Texto Corto

Estado actual del workflow

fechaCreacion

Fecha/Hora

Timestamp de creación

usuarioCreacion

Texto Corto

Email del usuario creador

fechaPaseTecnico

Fecha/Hora

Timestamp de cuando Calidad envía la solicitud a un Técnico

fechaCompletadoTecnico

Fecha/Hora

Timestamp de cuando el Técnico finaliza su parte

8.2. tbDatosPC (Propuesta de Cambio - F4203.11)
Campo

Tipo de Dato

idDatosPC

Autonumérico

idSolicitud

Numérico

refContratoInspeccionOficial

Texto Corto

refSuministrador

Texto Corto

suministradorNombreDir

Memo

objetoContrato

Memo

descripcionMaterialAfectado

Memo

numPlanoEspecificacion

Texto Corto

descripcionPropuestaCambio

Memo

descripcionPropuestaCambioCont

Memo

motivoCorregirDeficiencias

Sí/No

motivoMejorarCapacidad

Sí/No

motivoAumentarNacionalizacion

Sí/No

motivoMejorarSeguridad

Sí/No

motivoMejorarFiabilidad

Sí/No

motivoMejorarCosteEficacia

Sí/No

motivoOtros

Sí/No

motivoOtrosDetalle

Texto Corto

incidenciaCoste

Texto Corto

incidenciaPlazo

Texto Corto

incidenciaSeguridad

Sí/No

incidenciaFiabilidad

Sí/No

incidenciaMantenibilidad

Sí/No

incidenciaIntercambiabilidad

Sí/No

incidenciaVidaUtilAlmacen

Sí/No

incidenciaFuncionamientoFuncion

Sí/No

cambioAfectaMaterialEntregado

Sí/No

cambioAfectaMaterialPorEntregar

Sí/No

firmaOficinaTecnicaNombre

Texto Corto

firmaRepSuministradorNombre

Texto Corto

observacionesRACRef

Texto Corto

racCodigo

Texto Corto

observacionesRAC

Memo

fechaFirmaRAC

Fecha/Hora

obsAprobacionAutoridadDiseno

Memo

firmaAutoridadDisenoNombreCargo

Texto Corto

fechaFirmaAutoridadDiseno

Fecha/Hora

decisionFinal

Texto Corto

obsDecisionFinal

Memo

cargoFirmanteFinal

Texto Corto

fechaFirmaDecisionFinal

Fecha/Hora

8.3. tbDatosCDCA (Desviación / Concesión - F4203.10)
Campo

Tipo de Dato

idDatosCDCA

Autonumérico

idSolicitud

Numérico

refSuministrador

Texto Corto

numContrato

Texto Corto

identificacionMaterial

Memo

numPlanoEspecificacion

Texto Corto

cantidadPeriodo

Texto Corto

numSerieLote

Texto Corto

descripcionImpactoNC

Memo

descripcionImpactoNCCont

Memo

refDesviacionesPrevias

Texto Corto

causaNC

Memo

impactoCoste

Texto Corto

clasificacionNC

Texto Corto

requiereModificacionContrato

Sí/No

efectoFechaEntrega

Memo

identificacionAutoridadDiseno

Texto Corto

esSuministradorAD

Sí/No

racRef

Texto Corto

racCodigo

Texto Corto

observacionesRAC

Memo

fechaFirmaRAC

Fecha/Hora

decisionFinal

Texto Corto

observacionesFinales

Memo

fechaFirmaDecisionFinal

Fecha/Hora

cargoFirmanteFinal

Texto Corto

8.4. tbDatosCDCASUB (Desviación / Concesión Sub-suministrador - F4203.101)
Campo

Tipo de Dato

idDatosCDCASUB

Autonumérico

idSolicitud

Numérico

refSuministrador

Texto Corto

refSubSuministrador

Texto Corto

suministradorPrincipalNombreDir

Memo

subSuministradorNombreDir

Memo

identificacionMaterial

Memo

numPlanoEspecificacion

Texto Corto

cantidadPeriodo

Texto Corto

numSerieLote

Texto Corto

descripcionImpactoNC

Memo

descripcionImpactoNCCont

Memo

refDesviacionesPrevias

Texto Corto

causaNC

Memo

impactoCoste

Texto Corto

clasificacionNC

Texto Corto

afectaPrestaciones

Sí/No

afectaSeguridad

Sí/No

afectaFiabilidad

Sí/No

afectaVidaUtil

Sí/No

afectaMedioambiente

Sí/No

afectaIntercambiabilidad

Sí/No

afectaMantenibilidad

Sí/No

afectaApariencia

Sí/No

afectaOtros

Sí/No

requiereModificacionContrato

Sí/No

efectoFechaEntrega

Memo

identificacionAutoridadDiseno

Texto Corto

esSubSuministradorAD

Sí/No

nombreRepSubSuministrador

Texto Corto

racRef

Texto Corto

racCodigo

Texto Corto

observacionesRAC

Memo

fechaFirmaRAC

Fecha/Hora

decisionSuministradorPrincipal

Texto Corto

obsSuministradorPrincipal

Memo

fechaFirmaSuministradorPrincipal

Fecha/Hora

firmaSuministradorPrincipalNombreCargo

Texto Corto

obsRACDelegador

Memo

fechaFirmaRACDelegador

Fecha/Hora

8.5. tbMapeoCampos
Tabla pre-poblada que contiene el mapeo entre los campos de las tablas de datos y los marcadores en las plantillas Word.

| Campo            | Tipo de Dato  |
| :--------------- | :------------ |
| idMapeo          | Autonumérico |
| nombrePlantilla  | Texto Corto   |
| nombreCampoTabla | Texto Corto   |
| valorAsociado    | Texto Corto   |
| nombreCampoWord  | Texto Corto   |

8.6. Tablas de Soporte
tbLogCambios: Auditoría de cambios en el sistema.

tbLogErrores: Registro de errores de la aplicación.

tbAdjuntos: Gestión de ficheros adjuntos a las solicitudes.

tbEstados: Definición de los estados del workflow.

tbTransiciones: Reglas para las transiciones de estado permitidas.

9. Anexo B: Mapeo de Campos para Generación de Documentos
   9.1. Plantilla "PC" (F4203.11 - Propuesta de Cambio)
   NombrePlantilla

NombreCampoTabla (en tbDatosPC)

ValorAsociado

NombreCampoWord

"PC"

refContratoInspeccionOficial

NULL

Parte0_1

"PC"

refSuministrador

NULL

Parte0_2

"PC"

suministradorNombreDir

NULL

Parte1_1

"PC"

objetoContrato

NULL

Parte1_2

"PC"

descripcionMaterialAfectado

NULL

Parte1_3

"PC"

numPlanoEspecificacion

NULL

Parte1_4

"PC"

descripcionPropuestaCambio

NULL

Parte1_5

"PC"

descripcionPropuestaCambioCont

NULL

Parte1_5Cont

"PC"

motivoCorregirDeficiencias

True

Parte1_6_1

"PC"

motivoMejorarCapacidad

True

Parte1_6_2

"PC"

motivoAumentarNacionalizacion

True

Parte1_6_3

"PC"

motivoMejorarSeguridad

True

Parte1_6_4

"PC"

motivoMejorarFiabilidad

True

Parte1_6_5

"PC"

motivoMejorarCosteEficacia

True

Parte1_6_6

"PC"

motivoOtros

True

Parte1_6_7

"PC"

motivoOtrosDetalle

NULL

Parte1_6_8

"PC"

incidenciaCoste

"Aumentará"

Parte1_7a_1

"PC"

incidenciaCoste

"Disminuirá"

Parte1_7a_2

"PC"

incidenciaCoste

"No variará"

Parte1_7a_3

"PC"

incidenciaPlazo

"Aumentará"

Parte1_7b_1

"PC"

incidenciaPlazo

"Disminuirá"

Parte1_7b_2

"PC"

incidenciaPlazo

"No variará"

Parte1_7b_3

"PC"

incidenciaSeguridad

True

Parte1_7c_1

"PC"

incidenciaFiabilidad

True

Parte1_7c_2

"PC"

incidenciaMantenibilidad

True

Parte1_7c_3

"PC"

incidenciaIntercambiabilidad

True

Parte1_7c_4

"PC"

incidenciaVidaUtilAlmacen

True

Parte1_7c_5

"PC"

incidenciaFuncionamientoFuncion

True

Parte1_7c_6

"PC"

cambioAfectaMaterialEntregado

True

Parte1_9_1

"PC"

cambioAfectaMaterialPorEntregar

True

Parte1_9_2

"PC"

firmaOficinaTecnicaNombre

NULL

Parte1_10

"PC"

firmaRepSuministradorNombre

NULL

Parte1_11

"PC"

observacionesRACRef

NULL

Parte2_1

"PC"

racCodigo

NULL

Parte2_2

"PC"

observacionesRAC

NULL

Parte2_3

"PC"

fechaFirmaRAC

NULL

Parte2_4

"PC"

obsAprobacionAutoridadDiseno

NULL

Parte3_1

"PC"

firmaAutoridadDisenoNombreCargo

NULL

Parte3_2

"PC"

fechaFirmaAutoridadDiseno

NULL

Parte3_3

"PC"

decisionFinal

"APROBADO"

Parte3_2_1

"PC"

decisionFinal

"NO APROBADO"

Parte3_2_2

"PC"

obsDecisionFinal

NULL

Parte3_3_1

"PC"

cargoFirmanteFinal

NULL

Parte3_3_2

"PC"

fechaFirmaDecisionFinal

NULL

Parte3_3_3

9.2. Plantilla "CDCA" (F4203.10 - Desviación / Concesión)
NombrePlantilla

NombreCampoTabla (en tbDatosCDCA)

ValorAsociado

NombreCampoWord

"CDCA"

refSuministrador

NULL

Parte0_1

"CDCA"

numContrato

NULL

Parte1_2

"CDCA"

identificacionMaterial

NULL

Parte1_3

"CDCA"

numPlanoEspecificacion

NULL

Parte1_4

"CDCA"

cantidadPeriodo

NULL

Parte1_5a

"CDCA"

numSerieLote

NULL

Parte1_5b

"CDCA"

descripcionImpactoNC

NULL

Parte1_6

"CDCA"

refDesviacionesPrevias

NULL

Parte1_7

"CDCA"

causaNC

NULL

Parte1_8

"CDCA"

impactoCoste

"Increased / aumentado"

Parte1_9_1

"CDCA"

impactoCoste

"Decreased / disminuido"

Parte1_9_2

"CDCA"

impactoCoste

"Unchanged / sin cambio"

Parte1_9_3

"CDCA"

clasificacionNC

"Major / Mayor"

Parte1_10_1

"CDCA"

clasificacionNC

"Minor / Menor"

Parte1_10_2

"CDCA"

requiereModificacionContrato

True

Parte1_12_1

"CDCA"

efectoFechaEntrega

NULL

Parte1_13

"CDCA"

identificacionAutoridadDiseno

NULL

Parte1_14

"CDCA"

esSuministradorAD

True

Parte1_18_1

"CDCA"

esSuministradorAD

False

Parte1_18_2

"CDCA"

descripcionImpactoNCCont

NULL

Parte1_20

"CDCA"

racRef

NULL

Parte2_21_1

"CDCA"

racCodigo

NULL

Parte2_21_2

"CDCA"

observacionesRAC

NULL

Parte2_21_3

"CDCA"

fechaFirmaRAC

NULL

Parte2_22

"CDCA"

decisionFinal

"APROBADO"

Parte3_23_1

"CDCA"

decisionFinal

"NO APROBADO"

Parte3_23_2

"CDCA"

observacionesFinales

NULL

Parte3_24_1

"CDCA"

fechaFirmaDecisionFinal

NULL

Parte3_24_2

"CDCA"

cargoFirmanteFinal

NULL

Parte3_24_4

9.3. Plantilla "CDCASUB" (F4203.101 - Desviación / Concesión Sub-suministrador)
NombrePlantilla

NombreCampoTabla (en tbDatosCDCASUB)

ValorAsociado

NombreCampoWord

"CDCASUB"

refSuministrador

NULL

Parte0_1

"CDCASUB"

refSubSuministrador

NULL

Parte0_2

"CDCASUB"

suministradorPrincipalNombreDir

NULL

Parte1_1

"CDCASUB"

subSuministradorNombreDir

NULL

Parte1_2

"CDCASUB"

identificacionMaterial

NULL

Parte1_5

"CDCASUB"

numPlanoEspecificacion

NULL

Parte1_6

"CDCASUB"

cantidadPeriodo

NULL

Parte1_7a

"CDCASUB"

numSerieLote

NULL

Parte1_7b

"CDCASUB"

descripcionImpactoNC

NULL

Parte1_8

"CDCASUB"

refDesviacionesPrevias

NULL

Parte1_9

"CDCASUB"

causaNC

NULL

Parte1_10

"CDCASUB"

impactoCoste

"Incrementado"

Parte1_11_1

"CDCASUB"

impactoCoste

"Sin cambio"

Parte1_11_2

"CDCASUB"

impactoCoste

"Disminuido"

Parte1_11_3

"CDCASUB"

clasificacionNC

"Mayor"

Parte1_12_1

"CDCASUB"

clasificacionNC

"Menor"

Parte1_12_2

"CDCASUB"

afectaPrestaciones

True

Parte1_13_1

"CDCASUB"

afectaSeguridad

True

Parte1_13_2

"CDCASUB"

afectaFiabilidad

True

Parte1_13_3

"CDCASUB"

afectaVidaUtil

True

Parte1_13_4

"CDCASUB"

afectaMedioambiente

True

Parte1_13_5

"CDCASUB"

afectaIntercambiabilidad

True

Parte1_13_6

"CDCASUB"

afectaMantenibilidad

True

Parte1_13_7

"CDCASUB"

afectaApariencia

True

Parte1_13_8

"CDCASUB"

afectaOtros

True

Parte1_13_9

"CDCASUB"

requiereModificacionContrato

True

Parte1_14

"CDCASUB"

efectoFechaEntrega

NULL

Parte1_15

"CDCASUB"

identificacionAutoridadDiseno

NULL

Parte1_16

"CDCASUB"

esSubSuministradorAD

True

Parte1_20_1

"CDCASUB"

esSubSuministradorAD

False

Parte1_20_2

"CDCASUB"

nombreRepSubSuministrador

NULL

Parte1_21

"CDCASUB"

descripcionImpactoNCCont

NULL

Parte1_22

"CDCASUB"

racRef

NULL

Parte2_23_1

"CDCASUB"

racCodigo

NULL

Parte2_23_2

"CDCASUB"

observacionesRAC

NULL

Parte2_23_3

"CDCASUB"

fechaFirmaRAC

NULL

Parte2_25

"CDCASUB"

decisionSuministradorPrincipal

"APROBADO"

Parte3_26_1

"CDCASUB"

decisionSuministradorPrincipal

"NO APROBADO"

Parte3_26_2

"CDCASUB"

obsSuministradorPrincipal

NULL

Parte3_27_1

"CDCASUB"

fechaFirmaSuministradorPrincipal

NULL

Parte3_27_2

"CDCASUB"

firmaSuministradorPrincipalNombreCargo

NULL

Parte3_27_4

"CDCASUB"

obsRACDelegador

NULL

Parte4_28

"CDCASUB"

fechaFirmaRACDelegador

NULL

Parte4_30
