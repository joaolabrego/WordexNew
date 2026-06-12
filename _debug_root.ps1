$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $root = $wb.Worksheets.Item("ROOT")
    $lastCol = $root.Cells.Item(1, $root.Columns.Count).End(-4159).Column
    for ($col = 1; $col -le $lastCol; $col++) {
        $header = [string]$root.Cells.Item(1, $col).Text
        $cell = $root.Cells.Item(3, $col)
        $text = [string]$cell.Text
        if ($text -match "Erro") {
            Write-Output "ROOT col $col ($header): $($text.Substring(0, [Math]::Min(150, $text.Length)))"
        }
    }
    $json = $excel.Run("ObterRegistros", "ROOT")
    $idx = $json.IndexOf("Erro")
    if ($idx -ge 0) {
        Write-Output "JSON snippet: $($json.Substring([Math]::Max(0,$idx-80), 200))"
    }
    $wb.Close($false)
}
finally {
    $excel.Quit()
}
