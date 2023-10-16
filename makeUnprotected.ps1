# 현재 스크립트 파일의 경로 가져오기
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path $scriptPath

# 현재 스크립트 파일의 폴더를 기반으로 폴더 경로 설정
$folderPath = $scriptDirectory

# Excel COM 객체 생성
$excel = New-Object -ComObject Excel.Application

# 지정된 폴더 내의 모든 XLS 및 XLSX 파일 가져오기
$excelFiles = Get-ChildItem $folderPath | Where-Object { $_.Extension -eq ".xls" -or $_.Extension -eq ".xlsx" }

# 비밀번호를 빈 문자열로 변경할 때 사용할 새 비밀번호
$newPassword = ""

foreach ($file in $excelFiles) {
    # Excel 파일 열기 (비밀번호 입력)
    $workbook = $excel.Workbooks.Open($file.FullName)

    $workbook.Password = $newPassword
    # 비밀번호 변경
    $workbook.Unprotect()

    # 파일 저장 (비밀번호 해제)
    $workbook.Save()
    
    # 열려있는 Excel 워크북 닫기
    $workbook.Close()

    Write-Host "file $($file.Name). Unprotected"
}

# Excel 종료
$excel.Quit()

# COM 개체 해제
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

# PowerShell 세션에서 COM 개체 해제
Remove-Variable excel

Write-Host "Work Done!"
