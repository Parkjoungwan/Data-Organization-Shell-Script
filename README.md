# Data-Organization-Shell-Script
자료 정리용 쉘스크립트 임시저장용
각 PS1파일은 Window PowerShell 스크립트이며, 파일 이름 변환 및 엑셀 파일 수정 문서 작업을 위해 개인적으로 작성했습니다.
실행시키기 위해 권한설정이 필요할 수 있습니다.
각 파일의 작동은 다음과 같습니다.


-to_.ps1  
실행되는 폴더 내 파일 이름 내에 있는 "-" 문자를 "_" 문자로 치환합니다.  
makeUnprotected.ps1  
실행되는 폴더 내 엑셀 파일들의 비밀번호를 제거합니다. (비밀번호 입력이 필요합니다.)  
middleNameConvert.ps1  
middleNameConvert2.ps1  
실행되는 폴더 내 파일 이름을 특정 규칙에 맞춰 치환합니다.  
1. A_B_C_D_E의 C 내용을 입력값에 따라 치환합니다.  
2. A_B_C의 B 내용을 입력값에 따라 치환합니다.  
payrollRegisterConvert.ps1  
실해되는 폴더 내 파일이름 맨 앞에오는 날짜 문자열에서 연도, 월, 일로 잘라내 연도+월은 파일 맨 앞에 연도+월+일은 맨 뒤에 가도록 이름을 수정합니다.  
xls_to_xlsx.ps1  
실행되는 폴더 내 xls 파일을 xlsx 파일로 변환합니다.  
