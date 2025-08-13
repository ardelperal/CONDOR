Especificación Funcional y Arquitectura: Aplicación CONDOR
1. Resumen General del Proyecto
Nombre de la Aplicación: CONDOR

Contexto: Durante el ciclo de vida de un Expediente (el contrato principal adjudicado por la administración), pueden surgir necesidades de cambio o desviación sobre las características técnicas definidas. CONDOR es la herramienta que gestiona el ciclo de vida de estas peticiones.

Objetivo: Desarrollar una aplicación en Microsoft Access con VBA para que los usuarios con roles de Calidad y Técnico gestionen de forma colaborativa y centralizada el ciclo de vida completo de cada Solicitud (de No Conformidad, Desviación, Concesión o Cambio), de acuerdo con el procedimiento PA/42/03.

2. Arquitectura y Entorno de Ejecución
2.1. Despliegue y Actualización de Versiones
La aplicación sigue el modelo estándar de despliegue de la oficina, gestionado por una Lanzadera central.

Repositorio Central: Existe un directorio en la red (\\servidor\Aplicaciones\) que contiene la carpeta de cada aplicación, incluyendo CONDOR.

Estructura de Carpetas: Dentro de \\...\CONDOR\, la estructura es la siguiente:

recursos\: Contiene la última versión compilada del front-end (CONDOR.accde), un icono y un archivo de configuración (CONDOR.ini) que especifica la versión actual (ej: VersionAplicacion=2025-001).

CONDOR_Datos.accdb: Es el back-end de producción, que contiene únicamente las tablas.

Ejecución Local: La Lanzadera se encarga de copiar el contenido de la carpeta recursos\ a la carpeta local del usuario (%appdata%\Aplicaciones DYSN\CONDOR\).

Mecanismo de Actualización: Antes de ejecutar la aplicación, la Lanzadera compara la versión del archivo .ini del servidor con la del archivo .ini local del usuario. Si la versión del servidor es más reciente, actualiza los archivos locales antes de lanzar el .accde. Si no, ejecuta la versión local existente.

2.2. Gestión de Entornos (Local vs. Oficina)
La aplicación debe ser capaz de funcionar en dos entornos distintos sin necesidad de cambiar el código:

Entorno de Oficina (Producción): El archivo .accde ejecutado por los usuarios finales. Debe conectarse al back-end de producción en la red.

Entorno Local (Desarrollo/Test): El archivo .accdb utilizado por los desarrolladores. Debe conectarse a una copia local del back-end con datos de prueba (ubicada en una subcarpeta dbs-local).

Para lograr esto, la aplicación implementará la siguiente lógica:

Detección del Entorno: Al arrancar, una función en VBA comprobará el nombre del archivo actual.

Si termina en .accdb, se considerará entorno Local.

Si termina en .accde, se considerará entorno Oficina.

Configuración de Conexión: Se creará una tabla local en el front-end (TbConfiguracion) para almacenar las cadenas de conexión a las bases de datos de back-end.

Entorno

RutaBackend

Local

.\dbs-local\CONDOR_Datos_TEST.accdb

Oficina

\\Servidor\Aplicaciones\CONDOR\CONDOR_Datos.accdb

Re-enlace Automático de Tablas: Un procedimiento de inicio (AutoExec) ejecutará un código VBA que:

Detectará el entorno actual.

Leerá la RutaBackend correspondiente de la tabla TbConfiguracion.

Recorrerá todas las tablas enlazadas de la aplicación y actualizará su propiedad Connect para que apunten a la base de datos correcta.

Este mecanismo asegura que el mismo front-end se conecte a la base de datos de test durante el desarrollo y a la de producción cuando es desplegado por la Lanzadera.

3. Gestión de Usuarios y Roles
3.1. Login y Autenticación
La gestión de usuarios de CONDOR está integrada en el sistema central de aplicaciones de la oficina.

Lanzadera Central: Los usuarios acceden a CONDOR a través de un panel de control general, tras haberse autenticado en el sistema.

Paso de Parámetros: La lanzadera abrirá la aplicación CONDOR pasándole como parámetro de línea de comandos (cmd) el correo electrónico del usuario que ha iniciado sesión.

Identificación del Rol: Al iniciarse, CONDOR utilizará el correo electrónico recibido para consultar las tablas centrales de gestión de usuarios (TbUsuariosAplicaciones, TbUsuariosAplicacionesPermisos) y determinar qué rol tiene asignado el usuario dentro de la aplicación.

3.2. Roles en CONDOR
Rol

Interacción con CONDOR

Responsabilidades Principales

Calidad

Usuario Directo

Inicia las Solicitudes, rellena los campos iniciales, la libera para el rol Técnico. Tras la revisión, gestiona toda la comunicación externa y cierra la Solicitud.

Técnico

Usuario Directo

Corresponde al rol de Ingeniería. Recibe la Solicitud de Calidad, revisa los datos y modifica/añade la información técnica de su competencia. Libera la Solicitud de vuelta a Calidad.

Administrador

Usuario Directo

Tiene control total sobre la aplicación. Su principal función será la configuración y mantenimiento, como la gestión de la tabla TbMapeo_Campos para asegurar la correcta correspondencia con las plantillas.

Suministrador / RAC / Órgano de Contratación

Actor Externo

Reciben los documentos Word por correo electrónico, cumplimentan la parte que les corresponde y los devuelven al usuario de Calidad para su procesamiento. No acceden a la aplicación.

4. Flujo de Trabajo (Workflow)
La lógica se divide en dos fases claras: una interna de preparación y una externa de aprobación.

4.1. Flujo Interno de Preparación
Inicio por Calidad:

Un usuario con rol Calidad selecciona un Expediente vivo de una lista que se carga desde la aplicación central.

CONDOR determina automáticamente si para ese Expediente se actúa como Suministrador Principal o Sub-Suministrador.

Calidad elige el tipo de solicitud: Desviación/Concesión (CD-CA) o Propuesta de Cambio (PC).

Se crea la Solicitud con estado "Borrador Calidad".

Cumplimentación por Calidad:

Calidad rellena sus campos y pulsa "Liberar a Técnico". El estado cambia a "Pendiente Revisión Técnico".

Revisión de Técnico:

El rol Técnico recibe la Solicitud, añade/modifica la información técnica y pulsa "Liberar a Calidad". El estado cambia a "Pendiente Gestión Calidad".

4.2. Flujo Externo de Aprobación
Generación y Envío:

Calidad recibe la Solicitud completada internamente y genera el documento Word.

Envía el documento al actor externo correspondiente y actualiza el estado (ej: "Pendiente Respuesta RAC").

Recepción y Actualización:

Al recibir el documento de vuelta, Calidad utiliza la función "Leer Documento" (para .docx) o transcribe los datos (desde .pdf) para actualizar la Solicitud en la base de datos.

Ciclo y Cierre:

El proceso se repite hasta obtener la decisión final.

Calidad adjunta el PDF final y cambia el estado a "Cerrado".

5. Arquitectura Orientada a Objetos
Para facilitar el desarrollo, el mantenimiento y la realización de tests unitarios, se propone una arquitectura basada en la separación de responsabilidades y el uso de interfaces.

5.1. Separación en Capas
Capa de Presentación (Formularios): Los formularios de Access. Su código VBA se limitará a gestionar eventos y delegar la lógica a la capa de negocio.

Capa de Negocio (Lógica): Módulos de clase que orquestan el flujo de trabajo (WorkflowManager).

Capa de Acceso a Datos (DAL): Módulos de clase dedicados a interactuar con las tablas (SolicitudRepository).

Capa de Servicios Externos: Módulos de clase para interactuar con sistemas externos (WordService).

5.2. Uso de Interfaces para Testing
Se definirán interfaces (módulos de clase sin código) para las capas de servicios, como IWordService. Las clases de negocio dependerán de estas interfaces, no de las clases concretas. Esto permitirá "inyectar" clases mock (falsas) durante los tests para simular el comportamiento de servicios externos sin ejecutarlos realmente, facilitando así los tests unitarios automáticos.

6. Estructura de Datos y Mapeo de Campos
(El resto de la especificación, con el detalle de las tablas TbExpedientes, TbSolicitudes, TbDatos_CD_CA, TbDatos_PC, TbDatos_CD_CA_SUB y TbMapeo_Campos, se mantiene como en la versión anterior).