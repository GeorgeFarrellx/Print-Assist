param(
    [string]$AppDir = $PSScriptRoot,
    [string]$ShortcutDir = $env:PRINT_ASSIST_SHORTCUT_DIR
)

$ErrorActionPreference = 'Stop'

function Get-AppSourceFiles([string]$Directory) {
    foreach ($name in @('main.py', 'requirements.txt')) {
        $path = Join-Path $Directory $name
        if (Test-Path -LiteralPath $path -PathType Leaf) { Get-Item -LiteralPath $path }
    }
    $package = Join-Path $Directory 'print_assist'
    Get-ChildItem -LiteralPath $package -Recurse -File -Filter '*.py'
    $assets = Join-Path $package 'assets'
    if (Test-Path -LiteralPath $assets -PathType Container) {
        Get-ChildItem -LiteralPath $assets -Recurse -File
    }
}

function Get-SourceHash([System.IO.FileInfo]$File) {
    $bytes = [System.IO.File]::ReadAllBytes($File.FullName)
    if ($File.Extension -in @('.py', '.txt')) {
        $text = [System.Text.Encoding]::UTF8.GetString($bytes).Replace("`r`n", "`n")
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Test-CurrentBuild([string]$Executable, [System.IO.FileInfo[]]$Sources) {
    $infoPath = Join-Path (Split-Path -Parent $Executable) 'print-assist-build.json'
    if (-not (Test-Path -LiteralPath $infoPath -PathType Leaf)) { return $false }
    try {
        $info = Get-Content -LiteralPath $infoPath -Raw | ConvertFrom-Json
        if ($info.schema_version -ne 1) { return $false }
        if ((Get-SourceHash (Get-Item -LiteralPath $Executable)) -ne $info.executable_sha256) { return $false }
        if (@($info.source_hashes.PSObject.Properties).Count -ne $Sources.Count) { return $false }
        foreach ($file in $Sources) {
            $relative = $file.FullName.Substring($AppDir.Length).TrimStart('\', '/').Replace('\', '/')
            $property = $info.source_hashes.PSObject.Properties[$relative]
            if ($null -eq $property -or $property.Value -ne (Get-SourceHash $file)) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

try {
    $AppDir = (Get-Item -LiteralPath $AppDir).FullName.TrimEnd('\', '/')
    $hasSources = (Test-Path -LiteralPath (Join-Path $AppDir 'main.py') -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $AppDir 'print_assist') -PathType Container)
    $sources = @()
    if ($hasSources) { $sources = @(Get-AppSourceFiles $AppDir) }

    # Prefer the output of the current build specification over legacy copies.
    $candidates = @('dist\PrintAssist\PrintAssist.exe', 'PrintAssist.exe', 'dist\PrintAssist.exe')
    $target = $null
    $foundExecutable = $false
    foreach ($candidate in $candidates) {
        $path = Join-Path $AppDir $candidate
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $foundExecutable = $true
        if ($hasSources -and -not (Test-CurrentBuild $path $sources)) { continue }
        $target = $path
        break
    }
    if ($null -eq $target) {
        if ($foundExecutable -and $hasSources) {
            throw 'The available executable does not match the current Print Assist app files. Rebuild first: .venv\Scripts\python.exe -m PyInstaller --noconfirm --clean PrintAssist.spec (or use your installed Python). Your existing shortcut has not been changed.'
        }
        throw 'Print Assist executable was not found. Build the Windows app first: python -m PyInstaller --noconfirm --clean PrintAssist.spec.'
    }

    if ([string]::IsNullOrWhiteSpace($ShortcutDir)) { $ShortcutDir = [Environment]::GetFolderPath('Desktop') }
    if (-not (Test-Path -LiteralPath $ShortcutDir -PathType Container)) {
        New-Item -ItemType Directory -Path $ShortcutDir | Out-Null
    }
    $shortcut = Join-Path $ShortcutDir 'Print Assist.lnk'
    $icon = Join-Path $AppDir 'print_assist\assets\print-assist.ico'
    if (-not (Test-Path -LiteralPath $icon -PathType Leaf)) { $icon = $target }
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($shortcut)
    $link.TargetPath = $target
    $link.WorkingDirectory = Split-Path -Parent $target
    $link.IconLocation = $icon
    $link.Description = 'Open Print Assist'
    $link.Save()
    Write-Host "Created shortcut: $shortcut"
    Write-Host "Opens: $target"
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
