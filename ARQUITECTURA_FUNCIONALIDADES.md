# 🏗️ Arquitectura de Funcionalidades - Proyecto CONDOR

## 📋 Resumen de Componentes por Funcionalidad

### 🔐 **Autenticación (Auth)**
```
┌─────────────────────────────────────────────────────────────┐
│                    AUTENTICACIÓN                           │
├─────────────────────────────────────────────────────────────┤
│ 📄 IAuthService.cls          ← Interface                   │
│ 📄 IAuthRepository.cls       ← Interface                   │
│ 🔧 CAuthService.cls          ← Implementación              │
│ 🔧 CAuthRepository.cls       ← Implementación              │
│ 🧪 CMockAuthService.cls      ← Mock para testing           │
│ 🧪 CMockAuthRepository.cls   ← Mock para testing           │
│ 🏭 modAuthFactory.bas        ← Factory                     │
│ ✅ Test_AuthService.bas      ← Tests unitarios             │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CAuthService ➜ IAuthRepository
- CAuthService ➜ IErrorHandlerService
- CAuthRepository ➜ IConfig
```

### 📄 **Gestión de Documentos (Document)**
```
┌─────────────────────────────────────────────────────────────┐
│                GESTIÓN DE DOCUMENTOS                       │
├─────────────────────────────────────────────────────────────┤
│ 📄 IDocumentService.cls      ← Interface                   │
│ 🔧 CDocumentService.cls      ← Implementación              │
│ 🧪 CMockDocumentService.cls  ← Mock para testing           │
│ 🏭 modDocumentServiceFactory.bas ← Factory                 │
│ ✅ Test_DocumentService.bas  ← Tests unitarios             │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CDocumentService ➜ IWordManager
- CDocumentService ➜ IFileSystem
- CDocumentService ➜ IErrorHandlerService
```

### 📂 **Gestión de Expedientes (Expediente)**
```
┌─────────────────────────────────────────────────────────────┐
│                GESTIÓN DE EXPEDIENTES                      │
├─────────────────────────────────────────────────────────────┤
│ 📄 IExpedienteService.cls    ← Interface                   │
│ 📄 IExpedienteRepository.cls ← Interface                   │
│ 🔧 CExpedienteService.cls    ← Implementación              │
│ 🔧 CExpedienteRepository.cls ← Implementación              │
│ 🧪 CMockExpedienteRepository.cls ← Mock para testing       │
│ 🏭 modExpedienteServiceFactory.bas ← Factory               │
│ ✅ Test_CExpedienteService.bas ← Tests unitarios           │
│ 🔬 IntegrationTest_CExpedienteRepository.bas ← Tests integración │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CExpedienteService ➜ IExpedienteRepository
- CExpedienteService ➜ IErrorHandlerService
- CExpedienteRepository ➜ IConfig
```

### 📋 **Gestión de Solicitudes (Solicitud)**
```
┌─────────────────────────────────────────────────────────────┐
│                GESTIÓN DE SOLICITUDES                      │
├─────────────────────────────────────────────────────────────┤
│ 📄 ISolicitudService.cls     ← Interface                   │
│ 📄 ISolicitudRepository.cls  ← Interface                   │
│ 🔧 CSolicitudService.cls     ← Implementación              │
│ 🔧 CSolicitudRepository.cls  ← Implementación              │
│ 🧪 CMockSolicitudRepository.cls ← Mock para testing        │
│ 🏭 modSolicitudServiceFactory.bas ← Factory                │
│ ✅ Test_SolicitudService.bas ← Tests unitarios             │
│ 🔬 IntegrationTest_SolicitudRepository.bas ← Tests integración │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CSolicitudService ➜ ISolicitudRepository
- CSolicitudService ➜ IOperationLogger
- CSolicitudService ➜ IErrorHandlerService
```

### 🔄 **Gestión de Flujos de Trabajo (Workflow)**
```
┌─────────────────────────────────────────────────────────────┐
│              GESTIÓN DE FLUJOS DE TRABAJO                  │
├─────────────────────────────────────────────────────────────┤
│ 📄 IWorkflowService.cls      ← Interface                   │
│ 📄 IWorkflowRepository.cls   ← Interface                   │
│ 🔧 CWorkflowService.cls      ← Implementación              │
│ 🔧 CWorkflowRepository.cls   ← Implementación              │
│ 🧪 CMockWorkflowRepository.cls ← Mock para testing         │
│ 🏭 modWorkflowRepositoryFactory.bas ← Factory              │
│ ✅ Test_WorkflowService.bas  ← Tests unitarios             │
│ 🔬 IntegrationTest_WorkflowRepository.bas ← Tests integración │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CWorkflowService ➜ IWorkflowRepository
- CWorkflowService ➜ IErrorHandlerService
- CWorkflowRepository ➜ IConfig
```

### 🗺️ **Gestión de Mapeos (Mapeo)**
```
┌─────────────────────────────────────────────────────────────┐
│                  GESTIÓN DE MAPEOS                         │
├─────────────────────────────────────────────────────────────┤
│ 📄 IMapeoRepository.cls      ← Interface                   │
│ 🔧 CMapeoRepository.cls      ← Implementación              │
│ 🧪 CMockMapeoRepository.cls  ← Mock para testing           │
│ 🔬 IntegrationTest_CMapeoRepository.bas ← Tests integración │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CMapeoRepository ➜ IConfig
```

### 📧 **Gestión de Notificaciones (Notification)**
```
┌─────────────────────────────────────────────────────────────┐
│               GESTIÓN DE NOTIFICACIONES                    │
├─────────────────────────────────────────────────────────────┤
│ 📄 INotificationService.cls  ← Interface                   │
│ 📄 INotificationRepository.cls ← Interface                 │
│ 🔧 CNotificationService.cls  ← Implementación              │
│ 🧪 CMockNotificationService.cls ← Mock para testing        │
│ 🏭 modNotificationServiceFactory.bas ← Factory             │
│ ✅ Test_NotificationService.bas ← Tests unitarios          │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CNotificationService ➜ INotificationRepository
- CNotificationService ➜ IErrorHandlerService
```

### 📊 **Gestión de Operaciones y Logging (Operation)**
```
┌─────────────────────────────────────────────────────────────┐
│            GESTIÓN DE OPERACIONES Y LOGGING                │
├─────────────────────────────────────────────────────────────┤
│ 📄 IOperationLogger.cls      ← Interface                   │
│ 📄 IOperationRepository.cls  ← Interface                   │
│ 🔧 COperationLogger.cls      ← Implementación              │
│ 🔧 COperationRepository.cls  ← Implementación              │
│ 🧪 CMockOperationLogger.cls  ← Mock para testing           │
│ 🧪 CMockOperationRepository.cls ← Mock para testing        │
│ 🏭 modOperationLoggerFactory.bas ← Factory                 │
│ ✅ Test_OperationLogger.bas  ← Tests unitarios             │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- COperationLogger ➜ IOperationRepository
- COperationLogger ➜ IConfig
- COperationLogger ➜ IErrorHandlerService
```

### ⚙️ **Configuración (Config)**
```
┌─────────────────────────────────────────────────────────────┐
│                    CONFIGURACIÓN                           │
├─────────────────────────────────────────────────────────────┤
│ 📄 IConfig.cls               ← Interface                   │
│ 🔧 CConfig.cls               ← Implementación              │
│ 🧪 CMockConfig.cls           ← Mock para testing           │
│ 📦 modConfig.bas             ← Módulo auxiliar             │
│ 🔬 IntegrationTest_CConfig.bas ← Tests integración         │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CConfig ➜ IFileSystem (para leer archivos de configuración)
```

### 🗂️ **Sistema de Archivos (FileSystem)**
```
┌─────────────────────────────────────────────────────────────┐
│                 SISTEMA DE ARCHIVOS                        │
├─────────────────────────────────────────────────────────────┤
│ 📄 IFileSystem.cls           ← Interface                   │
│ 🔧 CFileSystem.cls           ← Implementación              │
│ 🧪 CMockFileSystem.cls       ← Mock para testing           │
│ 🧪 CMockTextFile.cls         ← Mock para archivos texto    │
│ 🏭 modFileSystemFactory.bas  ← Factory                     │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- Componente base sin dependencias externas
```

### 📝 **Gestión de Word (WordManager)**
```
┌─────────────────────────────────────────────────────────────┐
│                  GESTIÓN DE WORD                           │
├─────────────────────────────────────────────────────────────┤
│ 📄 IWordManager.cls          ← Interface                   │
│ 🔧 CWordManager.cls          ← Implementación              │
│ 🧪 CMockWordManager.cls      ← Mock para testing           │
│ 🏭 modWordManagerFactory.bas ← Factory                     │
│ ✅ Test_CWordManager.bas     ← Tests unitarios             │
│ 🔬 IntegrationTest_WordManager.bas ← Tests integración     │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- CWordManager ➜ IErrorHandlerService
- CWordManager ➜ IFileSystem
```

### ❌ **Gestión de Errores (ErrorHandler)**
```
┌─────────────────────────────────────────────────────────────┐
│                  GESTIÓN DE ERRORES                        │
├─────────────────────────────────────────────────────────────┤
│ 📄 IErrorHandlerService.cls  ← Interface                   │
│ 🔧 CErrorHandlerService.cls  ← Implementación              │
│ 🧪 CMockErrorHandlerService.cls ← Mock para testing        │
│ 🏭 modErrorHandlerFactory.bas ← Factory                    │
│ ✅ Test_ErrorHandlerService.bas ← Tests unitarios          │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- Componente base utilizado por casi todos los demás servicios
```

### 🧪 **Framework de Testing**
```
┌─────────────────────────────────────────────────────────────┐
│                  FRAMEWORK DE TESTING                      │
├─────────────────────────────────────────────────────────────┤
│ 📄 ITestReporter.cls         ← Interface                   │
│ 🔧 CTestReporter.cls         ← Implementación              │
│ 🔧 CTestResult.cls           ← Resultado de test           │
│ 🔧 CTestSuiteResult.cls      ← Resultado de suite          │
│ 📦 modAssert.bas             ← Utilidades de assert        │
│ 📦 modTestRunner.bas         ← Ejecutor de tests           │
│ ✅ Test_modAssert.bas        ← Tests del framework         │
└─────────────────────────────────────────────────────────────┘
```

### 🏢 **Gestión de Aplicación (AppManager)**
```
┌─────────────────────────────────────────────────────────────┐
│                GESTIÓN DE APLICACIÓN                       │
├─────────────────────────────────────────────────────────────┤
│ 📦 modAppManager.bas         ← Gestor principal            │
│ ✅ Test_AppManager.bas       ← Tests unitarios             │
└─────────────────────────────────────────────────────────────┘

🔗 **Dependencias:**
- modAppManager ➜ Múltiples servicios y factories
```

### 🗃️ **Modelos de Datos (Entidades)**
```
┌─────────────────────────────────────────────────────────────┐
│                   MODELOS DE DATOS                         │
├─────────────────────────────────────────────────────────────┤
│ 🏷️ T_Adjuntos.cls           ← Entidad Adjuntos            │
│ 🏷️ T_AuthData.cls           ← Entidad Datos Auth          │
│ 🏷️ T_Datos_CD_CA.cls        ← Entidad Datos CD/CA         │
│ 🏷️ T_Datos_CD_CA_SUB.cls    ← Entidad Datos CD/CA Sub     │
│ 🏷️ T_Datos_PC.cls           ← Entidad Datos PC            │
│ 🏷️ T_Estado.cls             ← Entidad Estado              │
│ 🏷️ T_Expediente.cls         ← Entidad Expediente          │
│ 🏷️ T_LogCambios.cls         ← Entidad Log Cambios         │
│ 🏷️ T_LogErrores.cls         ← Entidad Log Errores         │
│ 🏷️ T_Mapeo.cls              ← Entidad Mapeo               │
│ 🏷️ T_Operacion.cls          ← Entidad Operación           │
│ 🏷️ T_Solicitud.cls          ← Entidad Solicitud           │
│ 🏷️ T_Transicion.cls         ← Entidad Transición          │
│ 🏷️ T_Usuario.cls            ← Entidad Usuario             │
│ 🔧 QueryParameter.cls        ← Parámetro de consulta       │
└─────────────────────────────────────────────────────────────┘
```

### 📊 **Utilidades y Enumeraciones**
```
┌─────────────────────────────────────────────────────────────┐
│              UTILIDADES Y ENUMERACIONES                    │
├─────────────────────────────────────────────────────────────┤
│ 📦 modEnumeraciones.bas      ← Enumeraciones del sistema   │
│ 🏭 modRepositoryFactory.bas  ← Factory de repositorios     │
└─────────────────────────────────────────────────────────────┘
```

## 🔗 Mapa de Dependencias Principales

```
                    ┌─────────────────┐
                    │   IConfig       │
                    └─────────┬───────┘
                              │
              ┌───────────────┼───────────────┐
              │               │               │
    ┌─────────▼─────────┐    │    ┌─────────▼─────────┐
    │ IErrorHandler     │    │    │   IFileSystem     │
    │ Service           │    │    │                   │
    └─────────┬─────────┘    │    └─────────┬─────────┘
              │              │              │
              │              │              │
    ┌─────────▼─────────┐    │    ┌─────────▼─────────┐
    │   Todos los       │    │    │  CDocumentService │
    │   Servicios       │    │    │  CWordManager     │
    └───────────────────┘    │    └───────────────────┘
                             │
                   ┌─────────▼─────────┐
                   │  Repositorios     │
                   │  (Auth, Solicitud,│
                   │   Expediente,     │
                   │   Workflow, etc.) │
                   └───────────────────┘
```

## 📈 Estadísticas del Proyecto

- **Total de archivos:** 95
- **Interfaces:** 15
- **Implementaciones:** 15
- **Mocks:** 13
- **Factories:** 9
- **Tests unitarios:** 10
- **Tests de integración:** 6
- **Entidades de datos:** 14
- **Módulos auxiliares:** 6

## 🎯 Patrones Arquitectónicos Identificados

1. **Patrón Repository:** Separación entre lógica de negocio y acceso a datos
2. **Patrón Factory:** Creación centralizada de objetos
3. **Patrón Dependency Injection:** Inyección de dependencias a través de interfaces
4. **Patrón Mock Object:** Testing con objetos simulados
5. **Patrón Service Layer:** Capa de servicios para lógica de negocio

---
*Documento generado automáticamente - Proyecto CONDOR*