param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$VenvPython = Join-Path (Split-Path -Parent $Root) ".venv\Scripts\python.exe"
if (Test-Path $VenvPython) {
    $Python = $VenvPython
} else {
    $Python = "python"
}

Write-Host "Using Python: $Python"
if (-not $SkipInstall) {
    & $Python -m pip install --upgrade pip
    & $Python -m pip install -r requirements.txt
    & $Python -m pip install -r requirements-build.txt
}

Remove-Item -LiteralPath "build", "dist" -Recurse -Force -ErrorAction SilentlyContinue

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --name "work-scheduler-v3" `
    --add-data "web;web" `
    --add-data "templates;templates" `
    --collect-all "ortools" `
    --collect-all "holidays" `
    --collect-all "openpyxl" `
    --collect-submodules "uvicorn" `
    --exclude-module "pyarrow" `
    --exclude-module "matplotlib" `
    --exclude-module "PIL" `
    --hidden-import "uvicorn.logging" `
    --hidden-import "uvicorn.loops.auto" `
    --hidden-import "uvicorn.protocols.http.auto" `
    --hidden-import "uvicorn.protocols.websockets.auto" `
    --hidden-import "uvicorn.lifespan.on" `
    "launcher.py"

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code $LASTEXITCODE"
}

$DistDir = Join-Path $Root "dist\work-scheduler-v3"
if (-not (Test-Path $DistDir)) {
    throw "Build output folder was not created: $DistDir"
}

$RunBat = Join-Path $DistDir "run.bat"
@'
@echo off
cd /d "%~dp0"
start "" "%~dp0work-scheduler-v3.exe"
'@ | Set-Content -LiteralPath $RunBat -Encoding Default

$ReadmeTxt = Join-Path $DistDir "README.txt"
@'
근무표 생성기

근무 조건을 입력하면 월별 근무표를 자동으로 만들어 주는 프로그램입니다.

[실행 방법]
1. 이 폴더의 work-scheduler-v3.exe 또는 run.bat을 실행합니다.
2. 잠시 기다리면 브라우저가 자동으로 열립니다.
3. 브라우저에서 근무표를 작성합니다.

브라우저가 자동으로 열리지 않으면 실행 창에 표시된 주소를 직접 입력합니다.
기본 주소는 http://127.0.0.1:8007/ 입니다.
8007번 주소가 이미 사용 중이면 프로그램이 다른 번호의 주소를 자동으로 사용합니다.

[기본 사용 순서]
1. 왼쪽 설정에서 연도와 월을 선택합니다.
2. 필요한 경우 엑셀 서식을 불러옵니다.
3. 입력표에 직원 이름과 근무 조건을 입력합니다.
4. 날짜별로 미리 정해진 근무, 연가, 기타 근무를 입력합니다.
5. 근무표 생성 버튼을 누릅니다.
6. 결과를 확인합니다.
7. 필요하면 다시 생성 또는 부분 편집을 사용합니다.
8. 완료된 결과는 엑셀 다운로드로 저장합니다.

[종료 방법]
실행 중인 검은색 창을 닫거나 Ctrl+C를 누릅니다.
브라우저 탭만 닫으면 프로그램이 계속 실행 중일 수 있습니다.

[참고 사항]
- Python 설치는 필요하지 않습니다.
- 엑셀 다운로드 기능은 Microsoft Excel이 설치된 Windows 환경에서 사용하는 것을 권장합니다.
- 업로드한 엑셀 서식과 생성한 결과 파일은 사용자의 PC 안에서 처리됩니다.
'@ | Set-Content -LiteralPath $ReadmeTxt -Encoding UTF8

$ZipPath = Join-Path $Root "dist\work-scheduler-v3-windows.zip"
if (Test-Path $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}
Compress-Archive -LiteralPath $DistDir -DestinationPath $ZipPath -Force

Write-Host ""
Write-Host "Build complete:"
Write-Host $DistDir
Write-Host $ZipPath
