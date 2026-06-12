$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $ws = $wb.Worksheets.Item("Clientes")
    foreach ($addr in @("F3", "G3", "F4", "G4", "E3")) {
        $cell = $ws.Range($addr)
        Write-Output "--- $addr header=$($ws.Cells.Item(1, $cell.Column).Text) ---"
        Write-Output "Formula: $($cell.Formula)"
        Write-Output "Text: $($cell.Text.Substring(0, [Math]::Min(120, $cell.Text.Length)))"
    }
    $json = $excel.Run("ObterRegistros", "ROOT")
    if ($json -match "Erro") { Write-Output "JSON contem Erro" } else { Write-Output "JSON sem Erro" }
    $wb.Close($false)
}
finally {
    $excel.Quit()
}
