$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $excel.CalculateFull()
    $json = $excel.Run("ObterRegistros", "ROOT")
    if (-not $json) { throw "ObterRegistros retornou vazio." }
    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Close($false)
    Write-Output "OK: wordex.json regenerado a partir da macro atual da planilha."
}
catch {
    Write-Output "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    $excel.Quit()
}
