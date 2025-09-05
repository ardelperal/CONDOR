-- Limpiar datos existentes
DELETE FROM tbTransiciones;

-- Definir las transiciones de estado permitidas
-- De Borrador a Revisión Técnica (solo Calidad)
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (1, 1, 2, 'Calidad', TRUE);

-- De Revisión Técnica a Pendiente Aprobación (solo Técnico)
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (2, 2, 3, 'Técnico', TRUE);

-- De Pendiente Aprobación a Cerrado Aprobado (solo Calidad)
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (3, 3, 4, 'Calidad', TRUE);

-- De Pendiente Aprobación a Cerrado Rechazado (solo Calidad)
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (4, 3, 5, 'Calidad', TRUE);

-- Transiciones para el estado "En Tramitación" (ID 6)
-- 1. Desde "En Revisión Técnica" (ID 2) hacia "En Tramitación" (ID 6) - Acción del Técnico
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (5, 2, 6, 'Técnico', TRUE);

-- 2. Desde "En Tramitación" (ID 6) hacia "Cerrado - Aprobado" (ID 4) - Acción de Calidad
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (6, 6, 4, 'Calidad', TRUE);

-- 3. Desde "En Tramitación" (ID 6) hacia "Cerrado - Rechazado" (ID 5) - Acción de Calidad
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido, activa)
VALUES (7, 6, 5, 'Calidad', TRUE);