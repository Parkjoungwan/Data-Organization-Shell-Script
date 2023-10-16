# 사용자로부터 입력 받기
$clinicName = Read-Host "Please enter the hospital name:"

# 디렉토리 내의 파일 목록 가져오기
$files = Get-ChildItem

# 파일 목록 순회
foreach ($file in $files) {
    # 파일 이름과 확장자 분리
    $fileName = $file.Name
    $extension = $file.Extension

    # 파일 이름 분해
    $parts = $fileName -split "_"

    # 확장자 제거
    $nameWithoutExtension = $parts[4] -replace $extension, ""
    
    # 새 파일 이름 생성
    $newName = $parts[0] + "_" + $parts[1] + "_" + $clinicName + "_" + $parts[3] + "_" +$parts[4] + $extension
    
    # 새 파일 이름으로 파일 이름 변경
    Rename-Item -Path $file.FullName -NewName $newName
}
