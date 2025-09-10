-- Fichero: 002_schema_and_seed_tbTransiciones.sql
-- Eliminar registros existentes para idempotencia
DELETE FROM tbTransiciones;
-- Insertar las transiciones del flujo de trabajo
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (1, 1, 2, 'Calidad');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (2, 2, 3, 'Ingenieria');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (3, 3, 4, 'Calidad');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (4, 3, 5, 'Calidad');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (5, 4, 5, 'Calidad');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (6, 4, 5, 'Ingenieria');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (7, 5, 6, 'Calidad');
INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido)
VALUES (8, 6, 7, 'Calidad');