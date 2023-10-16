# 현재 스크립트 파일의 경로 가져오기
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = Split-Path $scriptPath

# 현재 스크립트 파일의 폴더를 기반으로 폴더 경로 설정
$folderPath = $scriptDirectory

# Excel COM 객체 생성
$excel = New-Object -ComObject Excel.Application

# 지정된 폴더 내의 모든 XLS 파일 가져오기
$xlsFiles = Get-ChildItem $folderPath -Filter *.xls

foreach ($file in $xlsFiles) {
    # Excel 파일 열기 (비밀번호 없이)
    $workbook = $excel.Workbooks.Open($file.FullName)

    # 새로운 파일 경로 생성 (확장자를 xlsx로 변경)
    $newPath = [System.IO.Path]::ChangeExtension($file.FullName, ".xlsx")

    # 파일 저장 (비밀번호 없이, 새 경로와 확장자를 xlsx로)
    $workbook.SaveAs($newPath, 51) # 51은 XLSX 형식을 나타냅니다

    # 열려있는 Excel 워크북 닫기
    $workbook.Close()

    Write-Host "file $($file.Name) changed xls to xlsx."
}

# Excel 종료
$excel.Quit()

# COM 개체 해제
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

# PowerShell 세션에서 COM 개체 해제
Remove-Variable excel

Write-Host "Work Done."
