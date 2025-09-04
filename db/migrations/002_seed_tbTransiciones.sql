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