$ErrorActionPreference = "Stop"

$repoSlug = if ($env:WDU_THESIS_SKILL_REPO) { $env:WDU_THESIS_SKILL_REPO } else { "CercaTrovato/wuhan-donghu-thesis-skill" }
$branch = if ($env:WDU_THESIS_SKILL_BRANCH) { $env:WDU_THESIS_SKILL_BRANCH } else { "main" }
$platform = if ($env:WDU_THESIS_SKILL_PLATFORM) { $env:WDU_THESIS_SKILL_PLATFORM.ToLowerInvariant() } else { "codex" }
$targetSkillsDir = if ($env:WDU_THESIS_SKILL_TARGET) { $env:WDU_THESIS_SKILL_TARGET } else { "" }
$archivePath = if ($env:WDU_THESIS_SKILL_ARCHIVE) { $env:WDU_THESIS_SKILL_ARCHIVE } else { "" }
$homeDir = [Environment]::GetFolderPath("UserProfile")

function Copy-WduSkillToSkillsDir {
  param(
    [Parameter(Mandatory = $true)][string]$SkillsDir,
    [Parameter(Mandatory = $true)][string]$SkillSource
  )

  if ([string]::IsNullOrWhiteSpace($SkillsDir)) {
    throw "Empty skills directory"
  }

  $destination = Join-Path $SkillsDir "wuhan-donghu-thesis"
  if ((Split-Path -Leaf $destination) -ne "wuhan-donghu-thesis") {
    throw "Refusing to install to unexpected path: $destination"
  }

  New-Item -ItemType Directory -Path $SkillsDir -Force | Out-Null
  if (Test-Path -LiteralPath $destination) {
    Remove-Item -LiteralPath $destination -Recurse -Force
  }
  Copy-Item -LiteralPath $SkillSource -Destination $destination -Recurse
  Write-Host "Installed wuhan-donghu-thesis -> $destination"
}

function Install-WduSkillForPlatform {
  param(
    [Parameter(Mandatory = $true)][string]$PlatformName,
    [Parameter(Mandatory = $true)][string]$SkillSource
  )

  switch ($PlatformName) {
    "codex" {
      $agentsHome = if ($env:AGENTS_HOME) { $env:AGENTS_HOME } else { Join-Path $homeDir ".agents" }
      $codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $homeDir ".codex" }
      Copy-WduSkillToSkillsDir -SkillsDir (Join-Path $agentsHome "skills") -SkillSource $SkillSource
      Copy-WduSkillToSkillsDir -SkillsDir (Join-Path $codexHome "skills") -SkillSource $SkillSource
    }
    "opencode" {
      $skillsDir = if ($env:OPENCODE_SKILLS_HOME) {
        $env:OPENCODE_SKILLS_HOME
      } elseif ($env:OPENCODE_CONFIG_DIR) {
        Join-Path $env:OPENCODE_CONFIG_DIR "skills"
      } else {
        Join-Path $homeDir ".config\opencode\skills"
      }
      Copy-WduSkillToSkillsDir -SkillsDir $skillsDir -SkillSource $SkillSource
    }
    "claude" {
      $claudeHome = if ($env:CLAUDE_HOME) { $env:CLAUDE_HOME } else { Join-Path $homeDir ".claude" }
      Copy-WduSkillToSkillsDir -SkillsDir (Join-Path $claudeHome "skills") -SkillSource $SkillSource
    }
    "cursor" {
      $cursorHome = if ($env:CURSOR_HOME) { $env:CURSOR_HOME } else { Join-Path $homeDir ".cursor" }
      Copy-WduSkillToSkillsDir -SkillsDir (Join-Path $cursorHome "skills") -SkillSource $SkillSource
    }
    "all" {
      Install-WduSkillForPlatform -PlatformName "codex" -SkillSource $SkillSource
      Install-WduSkillForPlatform -PlatformName "opencode" -SkillSource $SkillSource
      Install-WduSkillForPlatform -PlatformName "claude" -SkillSource $SkillSource
      Install-WduSkillForPlatform -PlatformName "cursor" -SkillSource $SkillSource
    }
    default {
      throw "Unknown WDU_THESIS_SKILL_PLATFORM '$PlatformName'. Use codex, opencode, claude, cursor, or all."
    }
  }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("wdu-thesis-skill-" + [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempRoot "repo.zip"
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

try {
  $zipUrl = "https://github.com/$repoSlug/archive/refs/heads/$branch.zip"
  if ($archivePath) {
    Write-Host "Using local archive $archivePath"
    Copy-Item -LiteralPath $archivePath -Destination $zipPath
  } else {
    Write-Host "Downloading $zipUrl"
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
  }
  Expand-Archive -LiteralPath $zipPath -DestinationPath $tempRoot -Force

  $repoRoot = Get-ChildItem -LiteralPath $tempRoot -Directory | Where-Object { $_.Name -ne "__MACOSX" } | Select-Object -First 1
  if (-not $repoRoot) {
    throw "Downloaded archive did not contain a repository directory"
  }

  $skillSource = Join-Path $repoRoot.FullName "skills\wuhan-donghu-thesis"
  if (-not (Test-Path -LiteralPath $skillSource)) {
    throw "Skill folder not found in archive: $skillSource"
  }

  if ($targetSkillsDir) {
    Copy-WduSkillToSkillsDir -SkillsDir $targetSkillsDir -SkillSource $skillSource
    return
  }

  Install-WduSkillForPlatform -PlatformName $platform -SkillSource $skillSource
} finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}
