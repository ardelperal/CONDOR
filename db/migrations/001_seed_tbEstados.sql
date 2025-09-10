 -- Fichero: 001_schema_and_seed_tbEstados.sql
-- Eliminar registros existentes para idempotencia
DELETE FROM tbEstados;
-- Insertar los estados del flujo de trabajo
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (1, 'Registrado', 'La solicitud ha sido registrada.', TRUE, FALSE, 10);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (2, 'Desarrollo', 'La solicitud está en fase de desarrollo técnico.', FALSE, FALSE, 20);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (3, 'Modificación', 'La solicitud requiere modificaciones de Calidad.', FALSE, FALSE, 30);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (4, 'Validación', 'La solicitud está pendiente de validación por RAC.', FALSE, FALSE, 40);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (5, 'Revisión', 'La solicitud está en revisión por el Cliente.', FALSE, FALSE, 50);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (6, 'Formalización', 'La solicitud está en fase de formalización final.', FALSE, FALSE, 60);
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (7, 'Aprobada', 'La solicitud ha sido aprobada y cerrada.', FALSE, TRUE, 70);