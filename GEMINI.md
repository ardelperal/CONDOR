# CONDOR Project Overview

## Project Summary
CONDOR is a Microsoft Access application developed with VBA, designed to manage the lifecycle of change requests, deviations, or concessions within public contract expedients. It is primarily used by Quality and Technical personnel.

## Architecture
The application features a centralized deployment via a launcher, with a clear separation between the front-end and back-end. It supports automatic version updates and can operate in both production (office) and local (development/test) modes without code changes.

## User Management
User login is integrated with a central system, and roles include Quality, Technical, Administrator, and external actors (who only receive documents).

## Workflow
The workflow is divided into an internal phase (preparation and review of requests by Quality and Technical teams) and an external phase (generation and sending of documents to external actors, reception, and closure).

## Code Architecture
The codebase is structured into distinct layers: Presentation, Business, Data Access, and External Services. This layered approach, along with the use of interfaces, facilitates unit testing.

## Data Structure
The core data structure includes main tables for Expedients, Solicitudes (Requests), specific data, and field mapping. The primary database file is `CONDOR_datos.accdb`.

## Building and Running
As a Microsoft Access application, there are no traditional build commands. The application is deployed centrally via a launcher. It runs directly from the Access environment, utilizing the `CONDOR_datos.accdb` file as its backend.

## Development Conventions
Development adheres to a layered architecture, promoting separation of concerns. The use of interfaces is encouraged to enable easier unit testing and maintainability.

## Key Files
- `back/CONDOR_datos.accdb`: The primary Microsoft Access database file containing the application's backend data.
- `docs/Plantillas/`: This directory contains various document templates (`.docx` files) used for generating official documents related to change requests, deviations, and concessions. Examples include `CD_CA_SUB`, `CD-CA`, and `PC` templates.
