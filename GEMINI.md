# CONDOR Project

## Project Overview

CONDOR is a Microsoft Access application written in VBA for managing the lifecycle of change, deviation, or concession requests related to public contracts. It is designed for Quality and Technical users.

The application follows a 3-tier architecture:

*   **Presentation Layer:** MS Access Forms
*   **Business Logic Layer:** VBA modules, including an `ExpedienteService` for interfacing with an existing system.
*   **Data Layer:** An MS Access database, which integrates with an existing records application via an `IDExpediente`.

## Building and Running

The project uses a VBScript command-line interface (`condor_cli.vbs`) to manage the development workflow. The VBA source code is maintained in the `src/` directory and then imported into the Access database.

**Key Commands:**

*   **Update modules in the database:**
    ```bash
    cscript condor_cli.vbs update
    ```
    This command synchronizes the files from the `src/` directory with the `CONDOR.accdb` database. It is the recommended way to update the database during development.

*   **Rebuild the entire project:**
    ```bash
    cscript condor_cli.vbs rebuild
    ```
    This command deletes all existing VBA modules from the database and re-imports them from the `src/` directory. This is a slower but more thorough way to ensure a clean build.

*   **Run tests:**
    To run the automated tests, open the `CONDOR.accdb` database, open the VBA editor (Alt+F11), and execute the `EJECUTAR_TODAS_LAS_PRUEBAS` subroutine from the `modAppManager` module. The test results will be displayed in the Immediate Window (Ctrl+G).

## Development Conventions

*   **Source Code:** All VBA code is stored as `.bas` and `.cls` files in the `src/` directory.
*   **Development Workflow:**
    1.  Modify the VBA code in the `src/` directory.
    2.  Use `condor_cli.vbs update` to import the changes into the `CONDOR.accdb` database.
    3.  Open `CONDOR.accdb` to manually test the changes and run the automated tests.
*   **Testing:** The project has an integrated testing framework.
    *   Test modules are named with the prefix `Test_`.
    *   The test runner is `modTestRunner.bas`.
    *   Tests are executed by running the `EJECUTAR_TODAS_LAS_PRUEBAS` subroutine from the `modAppManager` module.
*   **Configuration:** The application's configuration is managed in `modConfig.bas`. This module handles different environments (local vs. remote) and allows forcing a specific environment for testing purposes.

## Key Files and Directories

*   `condor_cli.vbs`: The command-line interface for managing the project.
*   `src/`: The directory containing all the VBA source code.
*   `back/Desarrollo/CONDOR.accdb`: The main development database.
*   `docs/`: Project documentation.
*   `docs/PLAN_DE_ACCION.md`: The project's roadmap and task list.
*   `README.md`: Detailed documentation about the project.
*   `GEMINI.md`: This file, providing context for the Gemini AI assistant.
