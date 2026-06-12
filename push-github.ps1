param(
  [string]$Message = "Update local installer $(Get-Date -Format 'yyyy-MM-dd HH:mm')",
  [string]$RepoName = "local-installer",
  [ValidateSet("private", "public")]
  [string]$Visibility = "public"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Run-Git {
  & git @args
  if ($LASTEXITCODE -ne 0) {
    throw "Git command failed: git $($args -join ' ')"
  }
}

function Run-Command {
  param(
    [string]$Exe,
    [string[]]$CommandArgs,
    [string]$FailureMessage
  )

  & $Exe @CommandArgs
  if ($LASTEXITCODE -ne 0) {
    throw $FailureMessage
  }
}

function Get-GhCommand {
  $gh = Get-Command gh -ErrorAction SilentlyContinue
  if ($gh) {
    return $gh.Source
  }

  $fallback = "D:\softs1\cmder\gh.cmd"
  if (Test-Path $fallback) {
    return $fallback
  }

  throw "GitHub CLI was not found. Install gh or add it to PATH."
}

if (-not (Test-Path ".git")) {
  Write-Host "Initializing git repository..."
  Run-Git init -b main
}

$ghExe = Get-GhCommand

Write-Host "Checking GitHub CLI authentication..."
& $ghExe auth status
if ($LASTEXITCODE -ne 0) {
  throw "GitHub CLI is not authenticated. Run: gh auth login"
}

Write-Host "Checking changes..."
Run-Git diff --check

if (Test-Path ".\generate_inventory.py") {
  Write-Host "Checking Python syntax..."
  Run-Command -Exe "python" -CommandArgs @("-m", "py_compile", ".\generate_inventory.py") -FailureMessage "Python syntax check failed."
}

Write-Host "Staging changes..."
Run-Git add --all -- .

& git diff --cached --quiet
if ($LASTEXITCODE -eq 1) {
  Write-Host "Creating commit: $Message"
  Run-Git commit -m $Message
} elseif ($LASTEXITCODE -ne 0) {
  throw "Unable to inspect staged changes."
} else {
  Write-Host "No new changes to commit."
}

$remotes = & git remote
if ($LASTEXITCODE -ne 0) {
  throw "Unable to inspect git remotes."
}
$hasOrigin = $remotes -contains "origin"

if (-not $hasOrigin) {
  Write-Host "Creating GitHub repository '$RepoName' as $Visibility..."
  & $ghExe repo create $RepoName "--$Visibility" --source . --remote origin
  if ($LASTEXITCODE -ne 0) {
    throw "GitHub repository creation failed."
  }
}

$proxy = "http://127.0.0.1:7897"
$githubUser = "bhcgdh"
$proxyListening = Get-NetTCPConnection -State Listen -LocalPort 7897 -ErrorAction SilentlyContinue
if (-not $proxyListening) {
  throw "Local GitHub proxy is not listening on 127.0.0.1:7897."
}

Write-Host "Pushing through local proxy..."
$env:GCM_INTERACTIVE = "Never"
Run-Git -c "credential.username=$githubUser" -c "credential.interactive=never" -c "http.proxy=$proxy" -c "https.proxy=$proxy" push -u origin main

Write-Host "Push completed."
Run-Git status -sb
Run-Git log -1 --oneline --decorate
