-- Fichero: 004_schema_tbConfiguracion.sql
-- Descripción: Define la estructura de la tabla de configuración.

CREATE TABLE tbConfiguracion (
    idConfiguracion COUNTER PRIMARY KEY,
    clave TEXT(255) NOT NULL,
    valor MEMO,
    descripcion TEXT(255),
    categoria TEXT(100),
    tipoValor TEXT(50),
    valorPorDefecto MEMO,
    esEditable YESNO,
    fechaCreacion DATETIME,
    fechaModificacion DATETIME,
    usuarioModificacion TEXT(100)
);