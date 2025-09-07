-- ============================================================================
-- Script: 002_seed_tbTransiciones.sql
-- Descripción: Datos semilla para tbTransiciones - Transiciones del flujo de trabajo
-- Fecha: 2024
-- Autor: CONDOR-Developer
-- ============================================================================

-- Eliminar tabla existente para garantizar idempotencia
DROP TABLE tbTransiciones;

-- Crear tabla tbTransiciones con nueva estructura
CREATE TABLE tbTransiciones (
    idTransicion LONG PRIMARY KEY,
    idEstadoOrigen LONG NOT NULL,
    idEstadoDestino LONG NOT NULL,
    rolRequerido TEXT(50) NOT NULL
);

-- Insertar las transiciones del nuevo flujo de trabajo
-- Basado en el flujo: Registrado -> Desarrollo -> Modificación -> Validación/Revisión -> Formalización -> Aprobada

INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES 
-- Registrado (1) -> Desarrollo (2) | Rol: Calidad
(1, 1, 2, 'Calidad'),

-- Desarrollo (2) -> Modificación (3) | Rol: Ingenieria
(2, 2, 3, 'Ingenieria'),

-- Modificación (3) -> Validación (4) | Rol: Calidad (Si hay cambios)
(3, 3, 4, 'Calidad'),

-- Modificación (3) -> Revisión (5) | Rol: Calidad (Si no hay cambios)
(4, 3, 5, 'Calidad'),

-- Validación (4) -> Revisión (5) | Rol: Calidad
(5, 4, 5, 'Calidad'),

-- Validación (4) -> Revisión (5) | Rol: Ingenieria
(6, 4, 5, 'Ingenieria'),

-- Revisión (5) -> Formalización (6) | Rol: Calidad
(7, 5, 6, 'Calidad'),

-- Formalización (6) -> Aprobada (7) | Rol: Calidad
(8, 6, 7, 'Calidad');

-- Verificación de la inserción
-- SELECT COUNT(*) AS TotalTransiciones FROM tbTransiciones;
-- SELECT * FROM tbTransiciones ORDER BY idTransicion;