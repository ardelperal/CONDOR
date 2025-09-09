    DROP TABLE tbEstados;
-- Crear tabla tbEstados si no existe
CREATE TABLE tbEstados (
    idEstado LONG PRIMARY KEY,
    nombreEstado TEXT(50) NOT NULL,
    descripcion TEXT(255),
    esEstadoInicial YESNO DEFAULT FALSE,
    esEstadoFinal YESNO DEFAULT FALSE,
    orden LONG
);