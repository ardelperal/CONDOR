-- Fichero: 003_configuracion_app.sql
-- Define el esquema y los datos semilla para la configuración de la aplicación.

DROP TABLE tbConfiguracion;

CREATE TABLE tbConfiguracion (
    idConfiguracion COUNTER PRIMARY KEY,
    clave TEXT(255) NOT NULL,
    valor MEMO,
    descripcion TEXT(255)
);

-- Parámetros de la Aplicación
INSERT INTO tbConfiguracion (clave, valor, descripcion) VALUES ('ID_APLICACION_CONDOR', '231', 'ID numérico de CONDOR en Lanzadera.');
INSERT INTO tbConfiguracion (clave, valor, descripcion) VALUES ('CORREO_ADMINISTRADOR', 'admin@condor.com', 'Correo del administrador del sistema.');

-- Nombres de Fichero de Plantillas
INSERT INTO tbConfiguracion (clave, valor, descripcion) VALUES ('TEMPLATE_NAME_PC', 'PC.docx', 'Nombre del fichero para Propuesta de Cambio.');
INSERT INTO tbConfiguracion (clave, valor, descripcion) VALUES ('TEMPLATE_NAME_CDCA', 'CD_CA.docx', 'Nombre del fichero para Concesión/Desviación.');
INSERT INTO tbConfiguracion (clave, valor, descripcion) VALUES ('TEMPLATE_NAME_CDCASUB', 'CD_CA_SUB.docx', 'Nombre del fichero para C/D de Sub-suministrador.');