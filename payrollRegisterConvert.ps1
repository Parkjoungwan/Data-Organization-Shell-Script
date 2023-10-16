# 디렉토리 내의 파일 목록 가져오기
$files = Get-ChildItem

# 파일 목록 순회
foreach ($file in $files) {
    # 파일 이름과 확장자 분리
    $fileName = $file.Name
    $extension = $file.Extension

    # 파일 이름 분해
    $parts = $fileName -split "_"
    
    # 년월과 일 분리
    $yyyymm = $parts[0].Substring(0, 6)
    $dd = $parts[0].Substring(6, 2)

    # 확장자 제거
    $nameWithoutExtension = $parts[3] -replace $extension, ""
    
    # 새 파일 이름 생성
    $newName = $yyyymm + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + "_" + $yyyymm + $dd + $extension
    
    # 새 파일 이름으로 파일 이름 변경
    Rename-Item -Path $file.FullName -NewName $newName
}

