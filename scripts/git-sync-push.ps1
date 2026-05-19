# Ежедневная синхронизация локального репозитория с GitHub.
# Коммит только при наличии изменений; push — всегда (если есть что отправить).
param(
  [string]$RepoRoot = $PSScriptRoot + '\..',
  [string]$CommitMessage = ''
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path $RepoRoot).Path
$logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ('git-sync-' + (Get-Date -Format 'yyyy-MM-dd') + '.log')

function Write-Log([string]$msg) {
  $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
  Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
}

try {
  Set-Location -LiteralPath $RepoRoot
  if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
    throw "Не найден .git в $RepoRoot"
  }

  $branch = (git rev-parse --abbrev-ref HEAD 2>$null).Trim()
  if (-not $branch) { $branch = 'main' }

  git fetch origin 2>&1 | Out-Null
  $status = git status --porcelain
  if ($status) {
    git add -A
    if (-not $CommitMessage) {
      $CommitMessage = 'auto: ежедневный снимок ' + (Get-Date -Format 'yyyy-MM-dd')
    }
    git commit -m $CommitMessage
    Write-Log "commit OK ($branch): $CommitMessage"
  } else {
    Write-Log "commit skip: нет изменений"
  }

  $ahead = git rev-list --count "origin/$branch..HEAD" 2>$null
  if ($LASTEXITCODE -ne 0) { $ahead = git rev-list --count '@{u}..HEAD' 2>$null }
  if ($ahead -and [int]$ahead -gt 0) {
    git push origin $branch
    Write-Log "push OK ($branch): $ahead commit(s)"
  } else {
    Write-Log 'push skip: нечего отправлять'
  }
} catch {
  Write-Log ('ERROR: ' + $_.Exception.Message)
  exit 1
}

exit 0
