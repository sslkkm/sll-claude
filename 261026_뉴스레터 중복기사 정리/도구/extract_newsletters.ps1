# 받은 편지함\뉴스레터 폴더에서 지난 실행 이후 수신분 본문을 텍스트로 뽑는다. 메일은 읽기만 한다.
# 사용: & extract_newsletters.ps1 -Out <저장경로> [-From 'yyyy-MM-dd HH:mm']
# 기간: 도구\last_run.txt 의 시각부터 지금까지. 파일이 없으면 직전 평일 09:00부터.
param([Parameter(Mandatory = $true)][string]$Out, [string]$From)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$stateFile = Join-Path $PSScriptRoot 'last_run.txt'
$now = Get-Date

if (-not $From) {
  if (Test-Path $stateFile) { $From = (Get-Content $stateFile -Raw).Trim() }
  else {
    $d = $now.Date.AddDays(-1)
    while ($d.DayOfWeek -in 'Saturday', 'Sunday') { $d = $d.AddDays(-1) }
    $From = $d.AddHours(9).ToString('yyyy-MM-dd HH:mm')
  }
}

# 클래식 Outlook이 꺼져 있으면 띄운다.
# 시작이 로그인 창·대화상자 때문에 오래 걸리는 날이 있어(11/03 45분) 고정 대기 대신 준비될 때까지 확인한다.
$launched = $false
if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
  Start-Process 'C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE'
  $launched = $true
  Start-Sleep 30
}
$waitMin = if ($launched) { 20 } else { 10 }
$deadline = (Get-Date).AddMinutes($waitMin)
$ol = $null
while (-not $ol) {
  try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $null = $ol.GetNamespace('MAPI').GetDefaultFolder(6)
  } catch {
    $ol = $null
    # 실행 중인 개체 표에 아직 없으면 COM 생성으로도 붙어 본다 (같은 권한이면 실행 중인 Outlook에 연결됨)
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
      try { $ol = New-Object -ComObject Outlook.Application; $null = $ol.GetNamespace('MAPI').GetDefaultFolder(6) } catch { $ol = $null }
    }
  }
  if ($ol) { break }
  if ((Get-Date) -ge $deadline) { break }
  Start-Sleep 30
}
if (-not $ol) {
  $proc = Get-Process OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $proc) { $why = '클래식 Outlook이 켜져 있지 않습니다. 새 Outlook만 켜져 있는지 확인하세요.' }
  else {
    $title = $proc.MainWindowTitle
    try { $null = $proc.Path; $elev = -not $proc.Path } catch { $elev = $true }
    if ($elev) { $why = 'Outlook이 관리자 권한으로 실행 중인 것 같습니다. Outlook을 닫고 일반 권한으로 다시 켜 주세요.' }
    elseif ($title -notlike '*Outlook*' -or $title -notlike '*-*') { $why = "Outlook이 ${waitMin}분 동안 시작을 마치지 못했습니다(창 제목: '$title'). 로그인 창이나 대화상자가 떠 있는지 확인해 주세요." }
    else { $why = "Outlook은 열려 있지만 ${waitMin}분 동안 연결이 거부됐습니다(창 제목: '$title'). 대화상자가 떠 있는지 확인해 주세요." }
  }
  throw "클래식 Outlook에 연결하지 못했습니다. $why"
}
$ns = $ol.GetNamespace('MAPI')
try { $ns.SendAndReceive($false) } catch {}
# 막 켠 경우 하위 폴더 동기화가 늦으므로 조금 더 기다린다
Start-Sleep $(if ($launched) { 180 } else { 60 })

$f = $ns.GetDefaultFolder(6).Folders.Item('뉴스레터')
$to = $now.ToString('yyyy-MM-dd HH:mm')
# Restrict는 한국어 로캘에서 24시간 표기를 잘못 읽으므로 AM/PM 형식으로 넘긴다
$inv = [Globalization.CultureInfo]::InvariantCulture
$filter = "[ReceivedTime] >= '" + ([datetime]::Parse($From)).ToString('MM/dd/yyyy hh:mm tt', $inv) + "' AND [ReceivedTime] < '" + $now.ToString('MM/dd/yyyy hh:mm tt', $inv) + "'"
$its = $f.Items.Restrict($filter)
$its.Sort('[ReceivedTime]')
$sb = New-Object Text.StringBuilder
[void]$sb.AppendLine("기간: $From ~ $to")
$n = 0
$skipped = 0
foreach ($m in $its) {
  # Teams 등 메신저 알림은 뉴스레터가 아니므로 제외
  try { $addr = $m.SenderEmailAddress } catch { $addr = '' }
  if ($addr -like '*teams.mail.microsoft*' -or $m.SenderName -like '*Teams에서*') { $skipped++; continue }
  $n++
  $body = $m.Body -replace '<https?://[^>]+>', '' -replace '(?m)^\s+$', '' -replace '(\r?\n){3,}', "`n`n"
  if ($body.Length -gt 12000) { $body = $body.Substring(0, 12000) + ' …(이하 생략)' }
  [void]$sb.AppendLine("===== [$n] $($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm')) | $($m.SenderName) <$($m.SenderEmailAddress)> | $($m.Subject)")
  [void]$sb.AppendLine($body)
}
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
[IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object Text.UTF8Encoding $false))

# 받은 편지함 전체의 같은 기간 수신 수도 남겨 미수신 여부를 판단할 수 있게 한다
$inboxCount = $ns.GetDefaultFolder(6).Items.Restrict($filter).Count
Set-Content -Path $stateFile -Value $to -Encoding UTF8
"기간 $From ~ $to | 뉴스레터 ${n}통 (Teams 알림 ${skipped}통 제외) | 받은편지함(하위폴더 제외) ${inboxCount}통 | 저장 $Out"
