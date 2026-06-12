$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    Write-Output "Modulos VBA:"
    foreach ($comp in $wb.VBProject.VBComponents) {
        Write-Output ("  " + $comp.Name + " type=" + $comp.Type)
    }
    Write-Output "Testando Import..."
    $r = $wb.VBProject.VBComponents.Import("d:\WordexNew\Wordex.bas")
    Write-Output ("Import result: " + ($r -ne $null))
    if ($r) { Write-Output ("Imported as: " + $r.Name) }
}
catch {
    Write-Output "ERRO: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
}
