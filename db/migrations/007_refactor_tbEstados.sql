-- ============================================================================
-- Script: 007_refactor_tbEstados.sql
-- Descripción: Refactorización completa de tbEstados para el nuevo flujo de trabajo
-- Fecha: 2024
-- Autor: CONDOR-Developer
-- ============================================================================

-- Eliminar tabla existente para garantizar idempotencia
DROP TABLE IF EXISTS tbEstados;

-- Crear tabla tbEstados con nueva estructura
CREATE TABLE tbEstados (
    idEstado LONG PRIMARY KEY,
    nombreEstado TEXT(100) NOT NULL,
    descripcion TEXT(255),
    esEstadoInicial YESNO DEFAULT FALSE,
    esEstadoFinal YESNO DEFAULT FALSE
);

-- Insertar los 7 nuevos estados del flujo de trabajo
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal) VALUES 
(1, 'Registrado', 'La solicitud ha sido registrada en el sistema y está lista para iniciar el proceso.', TRUE, FALSE),
(2, 'Desarrollo', 'La solicitud está siendo desarrollada por el equipo de ingeniería.', FALSE, FALSE),
(3, 'Modificación', 'La solicitud requiere modificaciones adicionales antes de continuar.', FALSE, FALSE),
(4, 'Validación', 'La solicitud está siendo validada por el equipo correspondiente.', FALSE, FALSE),
(5, 'Revisión', 'La solicitud está en proceso de revisión final.', FALSE, FALSE),
(6, 'Formalización', 'La solicitud está siendo formalizada para su aprobación final.', FALSE, FALSE),
(7, 'Aprobada', 'La solicitud ha sido aprobada y el proceso ha finalizado exitosamente.', FALSE, TRUE);

-- Verificación de la inserción
-- SELECT COUNT(*) AS TotalEstados FROM tbEstados;
-- SELECT * FROM tbEstados ORDER BY idEstado;