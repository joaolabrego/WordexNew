$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$path = "d:\WordexNew\wordex.xlsm"
$basWordex = "d:\WordexNew\Wordex.bas"
$basConsulta = "d:\WordexNew\WordexConsulta.bas"

try {
    Write-Output "Abrindo $path ..."
    $wb = $excel.Workbooks.Open($path)

    Write-Output "Acessando VBProject ..."
    $components = $wb.VBProject.VBComponents

    $toRemove = @()
    foreach ($comp in $components) {
        if ($comp.Name -eq "Wordex" -or $comp.Name -eq "WordexConsulta") {
            $toRemove += $comp
        }
    }

    Write-Output "Removendo $($toRemove.Count) modulos antigos ..."
    foreach ($comp in $toRemove) {
        $components.Remove($comp)
    }

    Write-Output "Importando Wordex.bas ..."
    $null = $components.Import($basWordex)
    Write-Output "Importando WordexConsulta.bas ..."
    $null = $components.Import($basConsulta)

    Write-Output "Calculando e gerando JSON ..."
    $excel.CalculateFull()
    $json = $excel.Run("ObterRegistros", "ROOT")
    if (-not $json) { throw "ObterRegistros retornou vazio." }
    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))

    Write-Output "Salvando planilha ..."
    $wb.Save()
    $wb.Close($true)

    Write-Output "OK: modulos reimportados, wordex.json regenerado, planilha salva."
}
catch {
    Write-Output "ERRO: $($_.Exception.Message)"
    Write-Output $_.ScriptStackTrace
    exit 1
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
