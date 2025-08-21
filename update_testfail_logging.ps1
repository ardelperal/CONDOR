# Script para actualizar todos los TestFail con logging centralizado
# Autor: CONDOR-Expert

$sourceDir = "c:\Proyectos\CONDOR\src"
$testFiles = @(
    "Test_CSolicitudService.bas",
    "Test_ErrorHandler.bas", 
    "Test_Database_Complete.bas",
    "Test_Config.bas",
    "Test_CConfig.bas",
    "Test_SolicitudFactory.bas",
    "Test_AppManager.bas",
    "Test_CAuthService.bas",
    "Test_CExpedienteService.bas",
    "Test_ErrorHandler_Extended.bas",
    "Test_CSolicitudPC.bas"
)

Write-Host "Iniciando actualizacion de logging centralizado en modulos de prueba..." -ForegroundColor Green

foreach ($file in $testFiles) {
    $filePath = Join-Path $sourceDir $file
    
    if (Test-Path $filePath) {
        Write-Host "Procesando: $file" -ForegroundColor Yellow
        
        # Leer contenido del archivo
        $content = Get-Content $filePath -Raw
        
        # Patron para encontrar TestFail sin logging
        $pattern = 'TestFail:\s*\r?\n\s*(Test_\w+)\s*=\s*False'
        
        # Reemplazar con logging centralizado
        $replacement = "TestFail:`r`n    modErrorHandler.LogError `"`$1`", Err.Number, Err.Description, `"$file`"`r`n    `$1 = False"
        
        $updatedContent = $content -replace $pattern, $replacement
        
        # Escribir contenido actualizado
        Set-Content -Path $filePath -Value $updatedContent -Encoding UTF8
        
        Write-Host "  Actualizado: $file" -ForegroundColor Green
    } else {
        Write-Host "  No encontrado: $file" -ForegroundColor Red
    }
}

Write-Host "Actualizacion completada." -ForegroundColor Green
Write-Host "Verificando archivos actualizados..." -ForegroundColor Yellow

# Verificar que los cambios se aplicaron
foreach ($file in $testFiles) {
    $filePath = Join-Path $sourceDir $file
    if (Test-Path $filePath) {
        $content = Get-Content $filePath -Raw
        $logErrorCount = ($content | Select-String -Pattern "modErrorHandler\.LogError" -AllMatches).Matches.Count
        Write-Host "${file}: $logErrorCount llamadas a modErrorHandler.LogError" -ForegroundColor Cyan
    }
}

Write-Host "Proceso completado exitosamente!" -ForegroundColor Green