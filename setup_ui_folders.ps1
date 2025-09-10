# Obtener la ruta del directorio raíz del proyecto (donde se ejecuta el script)
$projectRoot = Get-Location

# Definir las rutas de las nuevas carpetas
$uiPath = Join-Path -Path $projectRoot -ChildPath "ui"
$assetsPath = Join-Path -Path $uiPath -ChildPath "assets"
$definitionsPath = Join-Path -Path $uiPath -ChildPath "definitions"
$templatesPath = Join-Path -Path $uiPath -ChildPath "templates"

# Array de rutas a crear
$foldersToCreate = @($uiPath, $assetsPath, $definitionsPath, $templatesPath)

Write-Host "Verificando y creando la estructura de directorios para UI as Code..."

foreach ($folder in $foldersToCreate) {
    if (-not (Test-Path -Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
        Write-Host "Creado: $folder"
    } else {
        Write-Host "Ya existe: $folder"
    }
}

Write-Host "Estructura de directorios UI completada."