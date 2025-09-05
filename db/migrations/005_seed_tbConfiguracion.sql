-- Limpiar datos existentes para asegurar la idempotencia
DELETE FROM tbConfiguracion;

-- Insertar la configuración estructural y por defecto de la aplicación
INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('DATA_PATH', 'C:\Proyectos\CONDOR\back\CONDOR_datos.accdb', 'Ruta a la base de datos de datos principal de CONDOR.', 'Rutas', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('DATABASE_PASSWORD', '', 'Contraseña para la base de datos de CONDOR (si aplica). Vacío si no tiene.', 'Seguridad', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('ATTACHMENTS_PATH', 'C:\Proyectos\CONDOR\data\adjuntos', 'Directorio raíz donde se almacenan los ficheros adjuntos.', 'Rutas', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('TEMPLATES_PATH', 'C:\Proyectos\CONDOR\data\plantillas', 'Directorio donde se encuentran las plantillas Word.', 'Rutas', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('CORREOS_DB_PATH', 'C:\Proyectos\CONDOR\back\Correos_datos.accdb', 'Ruta a la base de datos de notificaciones por correo.', 'Rutas', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('LANZADERA_DATA_PATH', 'C:\Proyectos\CONDOR\back\Lanzadera_Datos.accdb', 'Ruta a la base de datos de Lanzadera para autenticación.', 'Rutas', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('LANZADERA_PASSWORD', 'dpddpd', 'Contraseña para la base de datos de Lanzadera.', 'Seguridad', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('ID_APLICACION_CONDOR', '231', 'ID numérico que identifica a CONDOR dentro del sistema Lanzadera.', 'Aplicación', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('CORREO_ADMINISTRADOR', 'admin@condor.com', 'Correo del administrador del sistema para recibir notificaciones críticas.', 'Notificaciones', Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, fechaCreacion)
VALUES ('LOG_FILE_PATH', 'condor_app.log', 'Ruta del fichero de log para errores y eventos.', 'Logging', Now());