# 스크립트가 실행되기 전에 파일 이름 형식 확인
$pattern = "^\d{8}_[A-Za-z]_[A-Za-z]_[A-Za-z]"
$files = Get-ChildItem

$invalidFiles = @()

foreach ($file in $files) {
    $fileName = $file.Name
    $parts = $fileName -split "_"
    
    # 파일 이름 형식 확인: yyyymmdd_A_B_C
    if ($parts.Count -eq 4 -and $parts[0] -match "^\d{8}") {
        # 파일 이름과 확장자 분리
        $extension = $file.Extension

        # Extract year-month and day
        $yyyymm = $parts[0].Substring(0, 6)
        $dd = $parts[0].Substring(6, 2)

        # Remove extension
        $nameWithoutExtension = $parts[3] -replace $extension, ""
        
        # Create a new file name
        $newName = $yyyymm + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + "_" + $yyyymm + $dd + $extension
        
        # Rename the file with the new name
        Rename-Item -Path $file.FullName -NewName $newName
    } else {
        # Track files with invalid name format
        $invalidFiles += $fileName
    }
}

if ($invalidFiles.Count -gt 0) {
    Write-Host "The following file names do not match the expected format:"
    $invalidFiles | ForEach-Object {
        Write-Host $_
    }
    Write-Host "Script execution skipped."
} else {
    Write-Host "Converted!"
}
