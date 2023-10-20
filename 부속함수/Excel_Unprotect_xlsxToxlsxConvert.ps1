# 1번 스크립트
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path $scriptPath
$folderPath = $scriptDirectory

$excel = New-Object -ComObject Excel.Application

$excelFiles = Get-ChildItem $folderPath | Where-Object { $_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx" }
$newPassword = ""

foreach ($file in $excelFiles) {
    $workbook = $excel.Workbooks.Open($file.FullName)

    $workbook.Password = $newPassword
    $workbook.Unprotect()

    $workbook.Save()
    $workbook.Close()

    Write-Host "$($file.Name) is Unprotected"
}

# 1번 스크립트 완료 후 Excel 객체 해제
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

Write-Host "Unprotect Done!"

# 2번 스크립트
$xlsFiles = Get-ChildItem $folderPath -Filter *.xls

# 동일한 Excel 객체 재사용
$excel = New-Object -ComObject Excel.Application

foreach ($file in $xlsFiles) {
    $workbook = $excel.Workbooks.Open($file.FullName)
    $newPath = [System.IO.Path]::ChangeExtension($file.FullName, ".xlsx")
    $workbook.SaveAs($newPath, 51)
    $workbook.Close()

    Write-Host "file $($file.Name) changed xls to xlsx."
}

# 2번 스크립트 완료 후 Excel 객체 해제
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel

Write-Host "Work Done!"
