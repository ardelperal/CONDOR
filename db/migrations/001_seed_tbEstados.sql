-- Limpiar datos existentes para asegurar la idempotencia
DELETE FROM tbEstados;

-- Insertar los estados estructurales del workflow
INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (1, 'Borrador', 'La solicitud ha sido creada pero no enviada a revisión técnica.', TRUE, FALSE, 10);

INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (2, 'En Revisión Técnica', 'La solicitud ha sido enviada al equipo técnico para su cumplimentación.', FALSE, FALSE, 20);

INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (3, 'Pendiente Aprobación Calidad', 'El equipo técnico ha completado su parte y la solicitud está lista para la gestión de Calidad.', FALSE, FALSE, 30);

INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (4, 'Cerrado - Aprobado', 'La solicitud ha sido aprobada y el ciclo ha finalizado.', FALSE, TRUE, 100);

INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (5, 'Cerrado - Rechazado', 'La solicitud ha sido rechazado y el ciclo ha finalizado.', FALSE, TRUE, 110);

INSERT INTO tbEstados (idEstado, nombreEstado, descripcion, esEstadoInicial, esEstadoFinal, orden)
VALUES (6, 'En Tramitación', 'La solicitud ha sido completada por el equipo técnico y está siendo gestionada por Calidad para su tramitación externa antes de la decisión final.', FALSE, FALSE, 40);