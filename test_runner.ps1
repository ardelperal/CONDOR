# Script temporal para ejecutar las pruebas del framework CONDOR
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase('c:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb')
    $result = $access.Run('modTestRunner.RunAllTests')
    Write-Output $result
} catch {
    Write-Output "Error: $($_.Exception.Message)"
} finally {
    if ($access) {
        $access.Quit()
    }
}