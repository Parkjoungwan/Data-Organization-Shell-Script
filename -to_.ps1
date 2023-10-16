# 디렉토리 내의 파일 목록 가져오기
$files = Get-ChildItem

# 파일 목록 순회
foreach ($file in $files) {
    # 파일 이름과 확장자 분리
    $fileName = $file.Name
    $extension = $file.Extension

    # 파일 이름 분해
    $parts = $fileName -split "-"
    
    # PARTS 배열의 길이 확인
    $partsCount = $parts.Length

    if ($partsCount -ge 3) {
        # 새 파일 이름 생성 (PARTS 배열의 길이에 따라 다르게 생성)
        if ($partsCount -eq 3) {
	    $nameWithoutExtension = $parts[2] -replace $extension, ""
            $newName = $parts[0] + "_" + $parts[1] + "_" + $nameWithoutExtension + $extension
        }
        elseif ($partsCount -eq 4) {
	    $nameWithoutExtension = $parts[3] -replace $extension, ""
            $newName = $parts[0] + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + $extension
        }
        elseif ($partsCount -eq 5) {
	    $nameWithoutExtension = $parts[4] -replace $extension, ""
            $newName = $parts[0] + "_" + $parts[1] + "_" + $parts[2] + "_" + $parts[3] + "_" + $nameWithoutExtension + $extension
        }
        else {
            # 다른 경우에 대한 처리를 추가할 수 있습니다.
        }

        # 새 파일 이름으로 파일 이름 변경
        Rename-Item -Path $file.FullName -NewName $newName
    }
}
