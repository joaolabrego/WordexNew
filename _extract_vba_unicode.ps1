function Extract-VbaModule($path, $moduleName, $outPath) {
    $bytes = [IO.File]::ReadAllBytes($path)
    $u = [Text.Encoding]::Unicode.GetString($bytes)
    $startPat = "Attribute VB_Name = `"$moduleName`""
    $start = $u.IndexOf($startPat)
    if ($start -lt 0) {
        Write-Output "Module $moduleName not found"
        return
    }
    $next = $u.IndexOf('Attribute VB_Name = "', $start + $startPat.Length)
    if ($next -lt 0) { $next = $u.Length }
    $chunk = $u.Substring($start, $next - $start)
    $chunk = $chunk -replace '[^\p{L}\p{N}\p{P}\p{Z}\r\n]', ''
    [IO.File]::WriteAllText($outPath, $chunk, [Text.UTF8Encoding]::new($false))
    Write-Output "Wrote $outPath ($($chunk.Length) chars)"
}

Extract-VbaModule "D:\WordexNew\_xlsm_extract\xl\vbaProject.bin" "Wordex" "D:\WordexNew\_extracted_Wordex.bas"
Extract-VbaModule "D:\WordexNew\_xlsm_extract\xl\vbaProject.bin" "WordexTotais" "D:\WordexNew\_extracted_WordexTotais.bas"
Extract-VbaModule "D:\WordexNew\_xlsm_extract\xl\vbaProject.bin" "WordexGraficos" "D:\WordexNew\_extracted_WordexGraficos.bas"
