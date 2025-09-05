-- Fichero: 001_seed_tbEstados.sql
-- Descripción: Define la estructura y datos de la tabla de estados con clave primaria explícita.

-- Eliminar la tabla si ya existe para asegurar una recreación limpia
DROP TABLE tbEstados;

-- Crear la tabla con la clave primaria controlada por nosotros
CREATE TABLE tbEstados (
    idEstado LONG NOT NULL PRIMARY KEY,
    nombreEstado TEXT(255),
    descripcion MEMO,
    esEstadoInicial YESNO,
    esEstadoFinal YESNO,
    orden LONG
);

-- Insertar los estados con IDs explícitos
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (1, 'Borrador', 'La solicitud ha sido creada pero no enviada.', TRUE, FALSE, 10);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (2, 'En Revisión Técnica', 'La solicitud ha sido enviada al equipo técnico para su evaluación.', FALSE, FALSE, 20);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (3, 'Pendiente Aprobación Calidad', 'La solicitud está pendiente de la aprobación final del equipo de Calidad.', FALSE, FALSE, 30);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (6, 'En Tramitación', 'La solicitud ha sido completada por el equipo técnico y está siendo gestionada por Calidad para su tramitación externa.', FALSE, FALSE, 40);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (4, 'Cerrado - Aprobado', 'La solicitud ha sido aprobada y el flujo de trabajo ha finalizado.', FALSE, TRUE, 100);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden) VALUES (5, 'Cerrado - Rechazado', 'La solicitud ha sido rechazada y el flujo de trabajo ha finalizado.', FALSE, TRUE, 110);