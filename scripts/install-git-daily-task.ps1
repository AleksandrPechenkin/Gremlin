# Регистрирует задачу Планировщика Windows: ежедневный push на GitHub.
param(
  [string]$Time = '20:00',
  [string]$TaskName = 'Gremlin-GitHub-Daily-Push'
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$syncScript = Join-Path $scriptDir 'git-sync-push.ps1'
if (-not (Test-Path $syncScript)) { throw "Не найден $syncScript" }

$repoRoot = (Resolve-Path (Join-Path $scriptDir '..')).Path
$psExe = (Get-Command powershell.exe).Source
$arg = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -RepoRoot "{1}"' -f $syncScript, $repoRoot
$action = New-ScheduledTaskAction -Execute $psExe -Argument $arg
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
Write-Host "Задача '$TaskName' создана: ежедневно в $Time"
Write-Host "Репозиторий: $repoRoot"
Write-Host "Логи: $(Join-Path $scriptDir 'logs')"
