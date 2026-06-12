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

    $colProdutos = Get-ColumnByHeader $ws $headerRow "Produtos"
    $colTotais = Get-ColumnByHeader $ws $headerRow "TotaisProdutos"
    $colGrafico = Get-ColumnByHeader $ws $headerRow "TotaisProdutosGrafico"

    if ($colProdutos -eq 0) { throw "Coluna Produtos nao encontrada." }
    if ($colTotais -eq 0) { throw "Coluna TotaisProdutos nao encontrada." }
    if ($colGrafico -eq 0) { throw "Coluna TotaisProdutosGrafico nao encontrada." }

    $ws.Cells.Item($typesRow, $colProdutos).Value2 = "collection"
    $ws.Cells.Item($typesRow, $colTotais).Value2 = "totals"
    $ws.Cells.Item($typesRow, $colGrafico).Value2 = "graph"

    $row = $detailsRow
    while ([string]$ws.Cells.Item($row, 1).Text -ne "") {
        $clienteIdCell = $ws.Cells.Item($row, 1).Address($false, $false)
        $ws.Cells.Item($row, $colProdutos).Formula = "=ObterRegistros(""Produtos"", ""ClienteId"", $clienteIdCell)"
        $ws.Cells.Item($row, $colTotais).Formula = "=ObterRegistroTotal(""Produtos"", """", ""Preço, Quantidade"", ""Sum"", ""ClienteId"", $clienteIdCell)"
        $ws.Cells.Item($row, $colGrafico).Formula = "=ObterGrafico(""Produtos"", """", ""Preço, Quantidade"", ""Sum"", ""ClienteId"", $clienteIdCell)"
        $row++
    }

    $excel.CalculateFull()

    $json = $excel.Run("ObterRegistros", "ROOT")
    if (-not $json) { throw "ObterRegistros retornou vazio." }

    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Save()
    $wb.Close($false)
    Write-Output "OK: formulas corrigidas e wordex.json regenerado."
}
catch {
    Write-Output "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
