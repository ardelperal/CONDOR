-- Fichero: 005_seed_tbConfiguracion.sql
-- Descripción: Inserta los datos de configuración iniciales para la aplicación.
-- Versión 2.0 - Corregida para coincidir con el esquema completo de la tabla.

-- Limpiar datos existentes para asegurar la idempotencia
DELETE FROM tbConfiguracion;

-- Insertar la configuración estructural y por defecto de la aplicación
INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('DATA_PATH', 'C:\Proyectos\CONDOR\back\CONDOR_datos.accdb', 'Ruta a la base de datos de datos principal.', 'Rutas', 'Texto', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('DATABASE_PASSWORD', '', 'Contraseña para la base de datos de CONDOR. Vacío si no tiene.', 'Seguridad', 'Texto', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('ATTACHMENTS_PATH', 'C:\Proyectos\CONDOR\data\adjuntos', 'Directorio raíz para ficheros adjuntos.', 'Rutas', 'Texto', TRUE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('TEMPLATES_PATH', 'C:\Proyectos\CONDOR\data\plantillas', 'Directorio de plantillas Word.', 'Rutas', 'Texto', TRUE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('CORREOS_DB_PATH', 'C:\Proyectos\CONDOR\back\Correos_datos.accdb', 'Ruta a la base de datos de notificaciones.', 'Rutas', 'Texto', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('LANZADERA_DATA_PATH', 'C:\Proyectos\CONDOR\back\Lanzadera_Datos.accdb', 'Ruta a la base de datos de Lanzadera.', 'Rutas', 'Texto', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('LANZADERA_PASSWORD', 'dpddpd', 'Contraseña para la base de datos de Lanzadera.', 'Seguridad', 'Texto', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('ID_APLICACION_CONDOR', '231', 'ID numérico de CONDOR en Lanzadera.', 'Aplicación', 'Numero', FALSE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('CORREO_ADMINISTRADOR', 'admin@condor.com', 'Correo del administrador del sistema.', 'Notificaciones', 'Texto', TRUE, Now());

INSERT INTO tbConfiguracion (clave, valor, descripcion, categoria, tipoValor, esEditable, fechaCreacion)
VALUES ('LOG_FILE_PATH', 'condor_app.log', 'Ruta del fichero de log para errores y eventos.', 'Logging', 'Texto', TRUE, Now());