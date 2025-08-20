# leeme

*Convertido automáticamente desde DOCX*

---

Explicación de todo el aplicativo

La aplicación comienza cuando el usuario abre la lanzadera de la oficina y elije CONDOR para abrir, en ese momento la lanzadera verifica en el servidor la versión que se está usando y mira en el directorio del usuario la versión cliente suya, si es la misma directamente abre la aplicación cliente del usuario metiendo por vba.command el correo electrónico del usuario, ya que se ha logeado en la lanzadera

Condor ha de tomar ese correo electrónico y mirando en la base de datos de la lanzadera haciendo una consulta con el IDAplicacion y el correo del usuario para ver el tipo de usuario que es
Todas las bases de datos en entorno de producción están en \\datoste\aplicaciones_dys\Aplicaciones PpD\ cada una en el nombre de la aplicación

Y en local en ./back

En el caso de la lanzadera , es donde están los usuarios , la tabla de aplicaciones y la tabla de permisos de aplicaciones

Nombre de la base de datos Lanzadera_Datos.accdb (siempre con contraseña todas las bases de datos,tanto locales como remotas)
Consulta para saber el rol del usuario que entra

Para saber si es Administrador se mira exclusivamente en la tabla de TbUsuariosAplicaciones (no depende de la aplicación) un administrador es administrador si está ahí
sql=”SELECT TbUsuariosAplicaciones.EsAdministrador

FROM TbUsuariosAplicaciones

WHERE (((TbUsuariosAplicaciones.CorreoUsuario)='CorreoUsuario

alvaro.gonzalezcaballero@telefonica.com'));”

El campo EsAdministrador puede ser ‘Sí’ o ‘No’

Si es ‘Sí’ ya no miramos para esa aplicación si es administrador o no

Para conocer si es de técnico de calidad o administrador (cuando no lo era a nivel global) se hace esta consulta

Sql=” SELECT TbUsuariosAplicacionesPermisos.EsUsuarioAdministrador, TbUsuariosAplicacionesPermisos.EsUsuarioCalidad, TbUsuariosAplicacionesPermisos. EsUsuarioTecnico

FROM TbUsuariosAplicacionesPermisos

WHERE (((TbUsuariosAplicacionesPermisos.CorreoUsuario)='beatriz.novalgutierrez@telefonica.com') AND ((TbUsuariosAplicacionesPermisos.IDAplicacion)=231));”

EsUsuarioAdministrador, EsUsuarioCalidad y EsUsuarioTecnico pueden ser ‘Sí’ o ‘No’

Estos son los únicos perfiles aceptados para CONDOR.

Parte de Expedientes

La base de datos se llama Expedientes_datos.accdb situada como el resto tanto en oficina como en local
Los tres tipos de solicitudes que se pueden hacer en condor siempre van a ser referidas a un expediente. O sea que expediente es el que va a sufrir concesiones/desviaciones y Propuestas de cambio. En CONDOR solo nos va a interesar el IDExpediente para relacionarlo con la tabla de expedientes y mantener la trazabilidad, pero lo primero que va a hacer un miembro de calidad va a ser elegir un expediente para el que le quiere hacer uno de los dos tipos de solicitudes CD-CA Y PC . Si es de subsuministrador o no ya lo determina la tabla de expedientes.
Para elegir un expediente el sistema ha de sacar una consulta de esa base de datos que nos dé los campos más significativos para que Calidad sepa cuál elegir

Sql=” SELECT TbExpedientes.IDExpediente, TbExpedientes.Nemotecnico, TbExpedientes.Titulo, TbExpedientes.CodExp, TbExpedientes.FechaInicioContrato, TbExpedientes.FechaFinContrato, TbExpedientes.FechaFinGarantia, TbUsuariosAplicaciones.Nombre AS ResponsableCalidad, TbExpedientes.ContratistaPrincipal

FROM TbExpedientes LEFT JOIN TbUsuariosAplicaciones ON TbExpedientes.IDResponsableCalidad = TbUsuariosAplicaciones.Id

WHERE (((TbExpedientes.Adjudicado)='Sí') AND ((TbExpedientes.Pecal)='Sí'))

ORDER BY TbExpedientes.IDExpediente DESC , TbExpedientes.ContratistaPrincipal DESC;”

ResponsableCalidad, es el nombre del responsable de calidad que lleva ese expediente, puede ser uno de los campos por los cuales filtrar en el formulario

En el ámbito militar, PECAL (Publicaciones Españolas de Calidad) hace referencia a un conjunto de normas que establecen los requisitos de aseguramiento de la calidad para empresas que suministran bienes y servicios al Ministerio de Defensa español. Estas normas son la adaptación nacional de las normas AQAP (Allied Quality Assurance Publications) de la OTAN.

En resumen, PECAL garantiza que las empresas proveedoras cumplan con estándares de calidad específicos para contratos con el Ministerio de Defensa, siguiendo las directrices de la OTAN

Luego los expedientes también tienen jefes de proyecto y responsables, que es otro de los campos por los que se pueden filtrar para encontrar uno de ellos, incluso cuando se entre con el rol de responsable técnico sería bueno que a cada uno le apareciera los CD-CA y PC que pertenezcan a los expediente en los que son Jefes de proyecto o responsables

Strsql=” SELECT TbExpedientes.IDExpediente, TbExpedientes.Nemotecnico, TbExpedientes.Titulo, TbExpedientes.CodExp, TbExpedientes.FechaInicioContrato, TbExpedientes.FechaFinContrato, TbExpedientes.FechaFinGarantia, TbExpedientes.ContratistaPrincipal, TbExpedientesResponsables.EsJefeProyecto, TbUsuariosAplicaciones.Nombre AS JP

FROM (TbExpedientes INNER JOIN TbExpedientesResponsables ON TbExpedientes.IDExpediente = TbExpedientesResponsables.IdExpediente) INNER JOIN TbUsuariosAplicaciones ON TbExpedientesResponsables.IdUsuario = TbUsuariosAplicaciones.Id

WHERE (((TbExpedientes.Adjudicado)='Sí') AND ((TbExpedientes.Pecal)='Sí') AND ((TbExpedientesResponsables.EsJefeProyecto)='Sí'))

ORDER BY TbExpedientes.IDExpediente DESC , TbExpedientes.ContratistaPrincipal DESC;”

Con respecto a expedientes no vamos a necesitar nada más, ni crear nuevos ni eliminarlos, solo listarlos y tomar de ellos el IDExpediente que va a la clave externa en todo esto

Objetivo real de la aplicación
La aplicación debe poder

Registrar CD-CA, CD-CA-SUB Y PC

Con los datos anteriores debe poder rellenar plantillas ya definidas para cada tipo de registro

Si la plantilla se rellena por fuera de la herramienta, la herramienta ha de ser capaz de poder actualizar las tablas con los datos que le vengan de la plantilla

Poder saber siempre en el estado en el que está una de estas solicitudes.

Flujo de trabajo de la aplicación

Solo va a haber tres roles en la herramienta

Calidad, Técnico y Administrador

## FLUJO DE TRABAJO

Calidad entra en alta de solicitud

Ha de elegir un expediente para la solicitud

Ha de elegir un suministrador de los que el expediente tiene registrados previamente (viene dado de la base de datos de expedientes)

Ha de elegir si es un CA-CD o un PC

Si elije un CA-CD. El sistema ya sabe  con el expediente elegido con el campo ContratistaPrincipal si es CA-CD O CA-CD-SUB
Si ContratistaPrincipal=’Sí’ significa que nosotros somos contratistas principales y por lo tanto la solicitud es CD-CA

Calidad rellena en el aplicativo los campos necesarios par comenzar la FASE DE REGISTRO. En ese momento el sistema ha de rellenar la fecha en que ha ocurrido esto, así ya se sabe que tiene que entrar un técnico a realizar su parte. En este momento para que el técnico se entere le tendría que llegar una notificación por correo electrónico. (ver apartado de como se generan estos avisos)

El técnico entra en el aplicativo

Va a su bandeja de entrada y ve aquellas solicitudes que tienen que ver con sus expedientes y puede ver cuáles están en la FASE TÉCNICA, las rellena y tiene que dar al botón de liberar y en ese momento se rellena la fecha “fechaCompletadoTecnico” con lo que se ha de registrar la notificación al miembro de calidad que comenzó el proceso.

Calidad ya está en la FASE DE DESARROLLO

Ahora va a ir mandando correos electrónicos por fuera del aplicativo para mandar la plantilla de Word rellena por CONDOR (se ha de poder guardar en el directorio de anexos cada versión que generamos desde el aplicativo) estos correos electrónicos no se van a registrar en CONDOR, con estos correos electrónicos los agentes externos van rellenando de la plantilla su parte y devuelven el correo. Cuando calidad recibe la plantilla rellena en parte la va registrando en el formulario correspondiente de Access y así continuamente hasta que se llegue al final de la solicitud con la concesión o denegación del CA-CD o PC

