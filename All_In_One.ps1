function Excel_Unprotect_And_xlsxConvert {
# 암호풀기
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

# xlsx 변환
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

}

function FileNameConvert_[A_B_Target_C_D] {
    # 2번 스크립트 내용
    Write-Host "Sample Pattern [A_B_<Target>_C_D]"
    $clinicName = Read-Host "Please input target name:"

    $files = Get-ChildItem

    foreach ($file in $files) {
        $fileName = $file.Name
        $extension = $file.Extension

        $parts = $fileName -split "_"

        $nameWithoutExtension = $parts[4] -replace $extension, ""

        $newName = $parts[0] + "_" + $parts[1] + "_" + $clinicName + "_" + $parts[3] + "_" +$parts[4] + $extension

        Rename-Item -Path $file.FullName -NewName $newName
    }

    Write-Host "Converted!"
}

function FileNameConvert_[A_Target_C] {
    # 3번 스크립트 내용
    Write-Host "Sample Pattern [A_<Target>_C]"
    $targetName = Read-Host "Please input TargetName:"

    $files = Get-ChildItem

    foreach ($file in $files) {
        $fileName = $file.Name
        $extension = $file.Extension

        $parts = $fileName -split "_"

        $nameWithoutExtension = $parts[2] -replace $extension, ""

        $newName = $parts[0] + "_" + $targetName + "_" + $nameWithoutExtension + $extension

        Rename-Item -Path $file.FullName -NewName $newName
    }

    Write-Host "Converted!"
}

function FileNameConvert_[yyyymmdd_A_B_C] {
    # 4번 스크립트 내용
    $pattern = "^\d{8}_[A-Za-z]_[A-Za-z]_[A-Za-z]"
    $files = Get-ChildItem
    $invalidFiles = @()

    foreach ($file in $files) {
        $fileName = $file.Name
        $parts = $fileName -split "_"

        if ($parts.Count -eq 4 -and $parts[0] -match "^\d{8}") {
            $extension = $file.Extension
            $yyyymm = $parts[0].Substring(0, 6)
            $dd = $parts[0].Substring(6, 2)
            $nameWithoutExtension = $parts[3] -replace $extension, ""
            $newName = $yyyymm + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + "_" + $yyyymm + $dd + $extension
            Rename-Item -Path $file.FullName -NewName $newName
        } else {
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
}

function FileNameConvert_[-to_] {
    # 5번 스크립트 내용
    $files = Get-ChildItem

    foreach ($file in $files) {
        $fileName = $file.Name
        $extension = $file.Extension
        $parts = $fileName -split "-"
        $partsCount = $parts.Length

        if ($partsCount -ge 3) {
            if ($partsCount -eq 3) {
                $nameWithoutExtension = $parts[2] -replace $extension, ""
                $newName = $parts[0] + "_" + $parts[1] + "_" + $nameWithoutExtension + $extension
            } elseif ($partsCount -eq 4) {
                $nameWithoutExtension = $parts[3] -replace $extension, ""
                $newName = $parts[0] + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + $extension
            } elseif ($partsCount -eq 5) {
                $nameWithoutExtension = $parts[4] -replace $extension, ""
                $newName = $parts[0] + "_" + $parts[1] + "_" + $parts[2] + "_" + $parts[3] + "_" + $nameWithoutExtension + $extension
            }
            Rename-Item -Path $file.FullName -NewName $newName
        }
    }

    Write-Host "Converted!"
}

function FileNameConvert_[A_B_C_D_LastDayOfMonth] {
    $files = Get-ChildItem

    foreach ($file in $files) {
        $fileName = $file.Name
        $extension = $file.Extension

        # 파일 이름을 _ 문자로 분할
        $parts = $fileName -split "_"

        # A 부분에서 yyyymm 형식의 날짜 정보 추출
        $dateInfo = $parts[0]

        if ($dateInfo -match '(\d{4})(\d{2})') {
            $year = $matches[1]
            $month = $matches[2]

            # 해당 월의 말일 계산
            $lastDayOfMonth = [System.DateTime]::DaysInMonth([int]$year, [int]$month)

            # yyyymmdd 형식으로 날짜를 만듭니다.
            $formattedDate = $year + $month + $lastDayOfMonth.ToString("00")
	$nameWithoutExtension = $parts[3] -replace $extension, ""

            $newName = $dateInfo + "_" + $parts[1] + "_" + $parts[2] + "_" + $nameWithoutExtension + "_" + $formattedDate + $extension

            Rename-Item -Path $file.FullName -NewName $newName
        }
    }

    Write-Host "Converted!"
}

function FileNameConvert_[A_B_C_D_Target] {
    Write-Host "Sample Pattern [A_B_C_D_<Target>]"
    $target = Read-Host "Please input target:"

    $files = Get-ChildItem

    foreach ($file in $files) {
        $fileName = $file.Name
        $extension = $file.Extension

        $parts = $fileName -split "_"

        $nameWithoutExtension = $parts[4] -replace $extension, ""

        $newName = $parts[0] + "_" + $parts[1] + "_" + $parts[2] + "_" + $parts[3] + "_" + $target + $extension

        Rename-Item -Path $file.FullName -NewName $newName
    }

    Write-Host "Converted!"
}


# 사용자로부터 어떤 스크립트를 실행할 것인지 입력 받기
$continue = $true
while ($continue) {
    Write-Host "Choose a script to run (Enter a number):"
    Write-Host "1. Excel Unprotect & XLSX Convert"
    Write-Host "2. File Name Convert [A_B_Target_C_D]"
    Write-Host "3. File Name Convert [A_Target_C]"
    Write-Host "4. File Name Convert [yyyymmdd_A_B_C]"
    Write-Host "5. File Name Convert [-to_]"
    Write-Host "6. File Name Convert [A_B_C_D_LastDayOfMonth]"
    Write-Host "7. File Name Convert [A_B_C_D_Target]"
    Write-Host "0. Quit"
    $choice = Read-Host "Please input Number:"

    switch ($choice) {
        '1' { Excel_Unprotect_And_xlsxConvert }
        '2' { FileNameConvert_[A_B_Target_C_D] }
        '3' { FileNameConvert_[A_Target_C] }
        '4' { FileNameConvert_[yyyymmdd_A_B_C] }
        '5' { FileNameConvert_[-to_] }
	'6' { FileNameConvert_[A_B_C_D_LastDayOfMonth] }
	'7' { FileNameConvert_[A_B_C_D_Target] }
        '0' { $continue = $false }
        default { Write-Host "Invalid choice. Please select a valid option." }
    }
}
