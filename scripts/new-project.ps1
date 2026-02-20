[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Name,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($Name -notmatch '^[a-z0-9]+(?:-[a-z0-9]+)*$') {
    throw "Invalid project name '$Name'. Use lowercase letters, numbers, and hyphens only (example: zendesk-ticket-sync)."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$templatePath = Join-Path $repoRoot 'projects/project-template'
$projectsPath = Join-Path $repoRoot 'projects'
$destinationPath = Join-Path $projectsPath $Name

if (-not (Test-Path $templatePath)) {
    throw "Template folder not found: $templatePath"
}

if (Test-Path $destinationPath) {
    if (-not $Force) {
        throw "Project already exists: $destinationPath. Re-run with -Force to overwrite."
    }

    Remove-Item -Path $destinationPath -Recurse -Force
}

Copy-Item -Path $templatePath -Destination $destinationPath -Recurse -Force

# Replace template placeholders in key text files.
$replaceFiles = @(
    (Join-Path $destinationPath 'README.md'),
    (Join-Path $destinationPath 'pyproject.toml'),
    (Join-Path $destinationPath 'config.example.yml'),
    (Join-Path $destinationPath 'src/main.py')
)

foreach ($file in $replaceFiles) {
    if (Test-Path $file) {
        $content = Get-Content -Path $file -Raw
        $updated = $content -replace 'project-template', $Name
        Set-Content -Path $file -Value $updated
    }
}

Write-Host "Created project: projects/$Name"
Write-Host "Next steps:"
Write-Host "1) cd projects/$Name"
Write-Host "2) Copy .env.example to .env"
Write-Host "3) python src/main.py"
Write-Host "4) python -m pytest"
