$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $ws = $wb.Worksheets.Item("Clientes")

    $headerRow = 1
    $typesRow = 2
    $detailsRow = 3

    $produtosCol = 0
    $lastCol = $ws.Cells.Item($headerRow, $ws.Columns.Count).End(-4159).Column

    for ($col = 1; $col -le $lastCol; $col++) {
        $title = [string]$ws.Cells.Item($headerRow, $col).Text
        if ($title -eq "Produtos") {
            $produtosCol = $col
            break
        }
    }

    if ($produtosCol -eq 0) {
        $totaisCol = 0
        for ($col = 1; $col -le $lastCol; $col++) {
            if ([string]$ws.Cells.Item($headerRow, $col).Text -eq "TotaisProdutos") {
                $totaisCol = $col
                break
            }
        }

        if ($totaisCol -gt 0) {
            $insertAt = $totaisCol
        } else {
            $insertAt = $lastCol + 1
        }

        $ws.Columns.Item($insertAt).Insert() | Out-Null
        $produtosCol = $insertAt
        $ws.Cells.Item($headerRow, $produtosCol).Value2 = "Produtos"
        $ws.Cells.Item($typesRow, $produtosCol).Value2 = "collection"
    } else {
        $ws.Cells.Item($typesRow, $produtosCol).Value2 = "collection"
    }

    $row = $detailsRow
    while ([string]$ws.Cells.Item($row, 1).Text -ne "") {
        $clienteIdCell = $ws.Cells.Item($row, 1).Address($false, $false)
        $formula = "=ObterRegistros(""Produtos"", ""ClienteId"", $clienteIdCell)"
        $ws.Cells.Item($row, $produtosCol).Formula = $formula
        $row++
    }

    $excel.CalculateFull()

    $json = $excel.Run("ObterRegistros", "ROOT")
    if (-not $json) { throw "ObterRegistros retornou vazio." }

    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Save()
    $wb.Close($false)
    Write-Output "OK: coluna Produtos adicionada e wordex.json regenerado."
}
catch {
    Write-Output "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
