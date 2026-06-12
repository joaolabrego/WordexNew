$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open("d:\WordexNew\wordex.xlsm")
    $excel.CalculateFullRebuild()
    Start-Sleep -Seconds 2
    $json = $excel.Run("ObterRegistros", "ROOT")
    [System.IO.File]::WriteAllText("d:\WordexNew\wordex.json", $json, (New-Object System.Text.UTF8Encoding $false))
    $wb.Save()
    $wb.Close($false)
    Write-Output "OK"
}
finally {
    $excel.Quit()
}
