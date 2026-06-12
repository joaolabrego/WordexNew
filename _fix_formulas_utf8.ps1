$preco = "Pre" + [char]0x00E7 + "o"
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
    $ws = $wb.Worksheets.Item("Clientes")

    $headerRow = 1
    $typesRow = 2
    $detailsRow = 3
    $cols = "Preço, Quantidade"

    $colProdutos = Get-ColumnByHeader $ws $headerRow "Produtos"
    $colTotais = Get-ColumnByHeader $ws $headerRow "TotaisProdutos"
    $colGrafico = Get-ColumnByHeader $ws $headerRow "TotaisProdutosGrafico"

    $row = $detailsRow
    while ([string]$ws.Cells.Item($row, 1).Text -ne "") {
        $clienteIdCell = $ws.Cells.Item($row, 1).Address($false, $false)
        $ws.Cells.Item($row, $colProdutos).Formula = "=ObterRegistros(""Produtos"", ""ClienteId"", $clienteIdCell)"
        $ws.Cells.Item($row, $colTotais).Formula = "=ObterRegistroTotal(""Produtos"", """", ""$cols"", ""Sum"", ""ClienteId"", $clienteIdCell)"
        $ws.Cells.Item($row, $colGrafico).Formula = "=ObterGrafico(""Produtos"", """", ""$cols"", ""Sum"", ""ClienteId"", $clienteIdCell)"
        $row++
    }

    $root = $wb.Worksheets.Item("ROOT")
    $colTotaisRoot = Get-ColumnByHeader $root $headerRow "TotaisProdutosGerais"
    $colGraficoRoot = Get-ColumnByHeader $root $headerRow "TotaisProdutosGrafico"
    if ($colTotaisRoot -gt 0) {
        $root.Cells.Item($detailsRow, $colTotaisRoot).Formula = "=ObterRegistroTotal(""Produtos"", """", ""$cols"", ""Sum"")"
    }
    if ($colGraficoRoot -gt 0) {
        $root.Cells.Item($detailsRow, $colGraficoRoot).Formula = "=ObterGrafico(""Produtos"", """", ""$cols"", ""Sum"")"
    }

    $excel.CalculateFull()

    $json = $excel.Run("ObterRegistros", "ROOT")
    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Save()
    $wb.Close($false)
    Write-Output "OK"
}
finally {
    $excel.Quit()
}
