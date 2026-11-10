# =========================================================
# Claude 가상머신 파일(vm_bundles)을 C에서 D로 옮기는 스크립트
#   1. 옮기기  : D로 복사 > 검증 > C 원본 이름 변경(보관) > 정션 연결
#   2. 되돌리기: 정션 제거 > C 원본 이름 복구
#   3. 원본 삭제: Cowork 정상 확인 후 C에 보관해 둔 원본 삭제
# Claude 앱을 완전히 종료한 뒤 같은 이름의 .cmd 파일을 더블클릭해 실행한다.
# =========================================================
param(
    [ValidateSet('', 'move', 'rollback', 'cleanup')][string]$Mode = '',
    [string]$Source = 'C:\Users\Jnote\AppData\Local\Claude-Data\claude_pzs8sxrjxfjjc\vm_bundles',
    [string]$TargetRoot = 'D:\개인\Claude VM',
    [switch]$TestMode   # 가짜 폴더 시험용: 관리자 권한, 앱 종료, 서비스 제어를 건너뜀
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$ServiceName = 'CoworkVMService'
$Target = Join-Path $TargetRoot 'vm_bundles'
$Old = $Source + '_C원본보관'
$script:ServiceStopped = $false

# 관리자 권한으로 다시 실행. G드라이브는 관리자 창에서 안 보일 수 있어 임시 폴더 사본으로 실행한다.
if (-not $TestMode) {
    $identity = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $temp = Join-Path $env:TEMP 'claude_vm_move.ps1'
        Copy-Item -LiteralPath $PSCommandPath -Destination $temp -Force
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$temp`""
        exit
    }
}

function Test-Link([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    return [bool]((Get-Item -LiteralPath $Path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)
}

function Stop-CoworkService {
    if ($TestMode) { return }
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -ne 'Stopped') {
        Write-Host "$ServiceName 서비스를 멈춥니다..."
        $script:ServiceStopped = $true
        Stop-Service -Name $ServiceName -Force
        (Get-Service -Name $ServiceName).WaitForStatus('Stopped', '00:00:30')
    }
}

function Start-CoworkService {
    if ($script:ServiceStopped) {
        Write-Host "$ServiceName 서비스를 다시 시작합니다..."
        Start-Service -Name $ServiceName
        $script:ServiceStopped = $false
    }
}

function Assert-NothingRunning {
    if ($TestMode) { return }
    # Claude 앱 본체만 확인한다. 종료되지 않고 남은 파일 검색 프로세스(claude.exe rg ...)는 vm_bundles를 쓰지 않는다.
    $app = Get-CimInstance Win32_Process -Filter "Name='claude.exe'" | Where-Object { ("$($_.ExecutablePath) $($_.CommandLine)") -match 'WindowsApps\\Claude_' }
    if ($app) {
        throw 'Claude가 아직 실행 중입니다. 작업 표시줄 오른쪽 트레이의 Claude 아이콘에서 종료한 뒤 다시 실행해 주세요.'
    }
    $copying = Get-CimInstance Win32_Process -Filter "Name='Robocopy.exe'" | Where-Object { $_.CommandLine -match 'Claude-Data' }
    if ($copying) {
        throw "Claude 백업 복사가 아직 진행 중입니다(프로세스 $($copying.ProcessId -join ', ')). 끝난 뒤 다시 실행해 주세요."
    }
}

function Assert-FilesNotInUse([string]$Path) {
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force)) {
        try {
            $stream = [IO.File]::Open($f.FullName, 'Open', 'Read', 'None')
            $stream.Close()
        } catch {
            throw "다른 프로그램이 사용 중인 파일이 있어 멈춥니다: $($f.Name). PC를 재시작한 뒤 Claude를 켜지 않고 다시 실행해 주세요."
        }
    }
}

function Assert-TreesMatch([string]$A, [string]$B) {
    $filesA = @(Get-ChildItem -LiteralPath $A -Recurse -File -Force)
    $filesB = @(Get-ChildItem -LiteralPath $B -Recurse -File -Force)
    if ($filesA.Count -ne $filesB.Count) { throw "파일 수가 다릅니다: 원본 $($filesA.Count)개, 사본 $($filesB.Count)개" }
    foreach ($f in $filesA) {
        $rel = $f.FullName.Substring($A.Length)
        $copy = Join-Path $B $rel.TrimStart('\')
        if (-not (Test-Path -LiteralPath $copy)) { throw "사본에 없는 파일: $rel" }
        if ((Get-Item -LiteralPath $copy -Force).Length -ne $f.Length) { throw "크기가 다른 파일: $rel" }
        if ((Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash) {
            throw "내용이 다른 파일: $rel"
        }
    }
    Write-Host "검증 완료: 파일 $($filesA.Count)개의 크기와 내용이 같습니다." -ForegroundColor Green
}

function Invoke-Move {
    Assert-NothingRunning
    if (-not (Test-Path -LiteralPath $Source)) { throw "원본 폴더가 없습니다: $Source" }
    if (Test-Link $Source) { throw '이미 정션으로 연결되어 있습니다. 옮기기는 끝난 상태입니다.' }
    if (Test-Path -LiteralPath $Old) { throw "보관 폴더가 이미 있습니다: $Old" }
    if (Test-Path -LiteralPath $Target) { throw "D드라이브에 같은 폴더가 이미 있습니다: $Target" }

    $size = (Get-ChildItem -LiteralPath $Source -Recurse -File -Force | Measure-Object Length -Sum).Sum
    $free = (New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($TargetRoot))).AvailableFreeSpace
    if ($free -lt ($size + 5GB)) { throw '대상 드라이브의 여유 공간이 부족합니다.' }

    Stop-CoworkService
    Assert-FilesNotInUse $Source

    $renamed = $false
    try {
        Write-Host ("{0:N2} GB를 복사합니다. 몇 분 걸릴 수 있습니다..." -f ($size / 1GB))
        & robocopy.exe $Source $Target /E /COPY:DATS /DCOPY:DAT /R:2 /W:5 /NFL /NDL /NP /NJH
        if ($LASTEXITCODE -ge 8) { throw "복사 실패 (robocopy 코드 $LASTEXITCODE)" }
        $global:LASTEXITCODE = 0

        Write-Host '복사본을 원본과 비교합니다...'
        Assert-TreesMatch $Source $Target

        Rename-Item -LiteralPath $Source -NewName (Split-Path $Old -Leaf)
        $renamed = $true
        New-Item -ItemType Junction -Path $Source -Target $Target | Out-Null

        if (-not (Test-Link $Source)) { throw '정션이 만들어지지 않았습니다.' }
        Assert-TreesMatch $Old $Source
    } catch {
        Write-Host '문제가 생겨 변경 전 상태로 되돌립니다...' -ForegroundColor Yellow
        if ($renamed) {
            if (Test-Link $Source) { [IO.Directory]::Delete($Source) }
            Rename-Item -LiteralPath $Old -NewName (Split-Path $Source -Leaf)
        }
        if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Recurse -Force }
        throw
    }

    Write-Host ''
    Write-Host '옮기기 완료.' -ForegroundColor Green
    Write-Host '다음 순서: Claude를 켜서 Cowork가 정상인지 확인한 뒤, 이 파일을 다시 실행해 3번(원본 삭제)을 고르세요.'
    Write-Host '문제가 있으면 Claude를 종료하고 2번(되돌리기)을 고르세요.'
}

function Invoke-Rollback {
    Assert-NothingRunning
    if (-not (Test-Link $Source)) { throw '정션 상태가 아니어서 되돌릴 것이 없습니다.' }
    if (-not (Test-Path -LiteralPath $Old)) { throw 'C드라이브 보관 폴더가 없어 되돌릴 수 없습니다. 원본 삭제가 이미 끝난 상태일 수 있습니다.' }

    Stop-CoworkService
    [IO.Directory]::Delete($Source)   # 정션 연결만 제거하며 D드라이브 파일은 건드리지 않음
    Rename-Item -LiteralPath $Old -NewName (Split-Path $Source -Leaf)

    Write-Host ''
    Write-Host '되돌리기 완료. C드라이브 원본을 다시 사용합니다.' -ForegroundColor Green
    Write-Host "D드라이브 사본은 남아 있습니다. 필요 없으면 직접 지우세요: $Target"
}

function Invoke-Cleanup {
    Assert-NothingRunning
    if (-not (Test-Link $Source)) { throw '정션 상태가 아니어서 원본을 지울 수 없습니다.' }
    $linked = ((Get-Item -LiteralPath $Source -Force).Target | Select-Object -First 1).TrimEnd('\')
    if ($linked -ne $Target.TrimEnd('\')) { throw "정션 대상이 예상과 다릅니다: $linked" }
    if (-not (Test-Path -LiteralPath $Old)) { throw '지울 보관 폴더가 없습니다.' }
    if (Test-Link $Old) { throw '보관 폴더가 정션입니다. 지우지 않습니다.' }
    # Cowork를 실행하면 D 쪽 파일 내용이 바뀌므로 내용 비교 대신 핵심 파일 존재만 확인한다
    if (-not (Test-Path -LiteralPath (Join-Path $Target 'claudevm.bundle\rootfs.vhdx'))) { throw 'D드라이브에서 핵심 파일(rootfs.vhdx)을 찾지 못했습니다. 원본을 지우지 않습니다.' }

    if (-not $TestMode) {
        Write-Host ''
        Write-Host "삭제할 폴더: $Old" -ForegroundColor Yellow
        Write-Host '삭제 후에는 2번(되돌리기)을 쓸 수 없습니다.' -ForegroundColor Yellow
        if ((Read-Host '삭제하려면 "삭제"라고 입력하세요') -ne '삭제') {
            Write-Host '취소했습니다.'
            return
        }
    }
    Remove-Item -LiteralPath $Old -Recurse -Force
    Write-Host '원본 삭제 완료. C드라이브 공간이 확보되었습니다.' -ForegroundColor Green
}

# ---- 실행 ----
if (-not $Mode) {
    Write-Host ''
    Write-Host 'Claude 가상머신 파일 이동' -ForegroundColor Cyan
    Write-Host "  현재 위치: $Source"
    Write-Host "  옮길 위치: $Target"
    Write-Host ''
    Write-Host '  1. 옮기기'
    Write-Host '  2. 되돌리기'
    Write-Host '  3. C드라이브 원본 삭제 (Cowork 정상 확인 후)'
    Write-Host ''
    switch (Read-Host '번호를 입력하세요') {
        '1' { $Mode = 'move' }
        '2' { $Mode = 'rollback' }
        '3' { $Mode = 'cleanup' }
    }
}

$transcribing = $false
try {
    if (-not $Mode) { throw '1, 2, 3 중에서 입력해 주세요. 아무것도 바꾸지 않았습니다.' }
    if (-not (Test-Path -LiteralPath $TargetRoot)) { New-Item -ItemType Directory -Path $TargetRoot | Out-Null }
    Start-Transcript -LiteralPath (Join-Path $TargetRoot '이동 로그.txt') -Append | Out-Null
    $transcribing = $true

    switch ($Mode) {
        'move'     { Invoke-Move }
        'rollback' { Invoke-Rollback }
        'cleanup'  { Invoke-Cleanup }
    }
} catch {
    Write-Host ''
    Write-Host "[중단] $($_.Exception.Message)" -ForegroundColor Red
    $global:LASTEXITCODE = 1
    $script:Failed = $true
} finally {
    try { Start-CoworkService } catch { Write-Host "[주의] 서비스를 다시 시작하지 못했습니다: $($_.Exception.Message) PC를 재시작해 주세요." -ForegroundColor Red }
    if ($transcribing) { try { Stop-Transcript | Out-Null } catch {} }
}

if (-not $TestMode) {
    Write-Host ''
    Read-Host '창을 닫으려면 Enter를 누르세요' | Out-Null
}
if ($script:Failed) { exit 1 }
