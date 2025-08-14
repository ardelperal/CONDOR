# Script para compilar proyecto de Access
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Access
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase('C:\Proyectos\CONDOR\CONDOR.accdb')
    
    Write-Host "Compilando módulos..."
    $access.DoCmd.RunCommand([Microsoft.Office.Interop.Access.AcCommand]::acCmdCompileAndSaveAllModules)
    
    Write-Host "Compilación completada exitosamente"
    $access.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
}
catch {
    Write-Host "Error durante la compilación: $($_.Exception.Message)"
    if ($access) {
        $access.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($access) | Out-Null
    }
    exit 1
}