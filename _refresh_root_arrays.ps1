$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Get-ColumnByHeader($ws, $headerRow, $name) {
    $lastCol = $ws.Cells.Item($headerRow, $ws.Columns.Count).End(-4159).Column
    for ($col = 1; $col -le $lastCol; $col++) {
        if ([string]$ws.Cells.Item($headerRow, $col).Text -eq $name) {
            return $col
        }
    }
    return 0
}

try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $root = $wb.Worksheets.Item("ROOT")

    $colClientes = Get-ColumnByHeader $root 1 "Clientes"
    if ($colClientes -gt 0) {
        $cell = $root.Cells.Item(3, $colClientes)
        $cell.Clear()
        $cell.Formula2 = "=ObterRegistros(""Clientes"")"
    }

    $excel.CalculateFullRebuild()
    Start-Sleep -Seconds 2

    $text = [string]$root.Cells.Item(3, $colClientes).Text
    Write-Output "Clientes cell has Produtos: $($text -match 'Produtos')"
    Write-Output "Clientes cell has Erro: $($text -match 'Erro')"

    $json = $excel.Run("ObterRegistros", "ROOT")
    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Save()
    $wb.Close($false)
    Write-Output "JSON Erro: $($json -match 'Erro')"
}
finally {
    $excel.Quit()
}
