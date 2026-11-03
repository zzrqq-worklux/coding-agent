# installer-core.ps1 - core logic for OpenCode 安装器 (v4)

$script:InstallerCoreVersion = '4.2.2'
$script:PowerShell7Version = '7.6.9'
$script:PowerShell7Sha256 = 'CE7484804C0F0AD4DDA10C958E71ACDAF7EEBFB22505C1B480F89E0E9F1DA94D'

function Get-EditionFromFile {
    param([string[]]$Paths)

    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $value = (Get-Content -LiteralPath $path -Raw -Encoding UTF8).Trim().ToLowerInvariant()
        } catch {
            continue
        }
        if ($value -eq 'portable' -or $value -eq 'full') { return $value }
    }
    return 'full'
}

function Get-EditionInfo {
    param([string]$Edition)

    if ($Edition -eq 'portable') {
        return @{
            Key         = 'portable'
            Name        = '便携版'
            Description = '仅内置 Outlook / Excel MCP，体积小'
            ExeName     = 'OpenCode安装器-便携版.exe'
        }
    }
    return @{
        Key         = 'full'
        Name        = '厚重版'
        Description = '内置 PPT-Master、归藏网页 PPT 与 Outlook / Excel MCP'
        ExeName     = 'OpenCode安装器-厚重版.exe'
    }
}

function Get-EnabledComponents {
    param([string]$Edition)

    if ((Get-EditionInfo -Edition $Edition).Key -eq 'portable') { return @('outlook', 'excel') }
    return @('ppt-master', 'guizang', 'outlook', 'excel')
}

function Get-RequiredPayloads {
    param([string]$Edition)

    $components = @(Get-EnabledComponents -Edition $Edition)
    $names = @()
    if ($components -contains 'ppt-master') { $names += 'ppt-master-skill-v6.7.3.zip' }
    if ($components -contains 'guizang') { $names += 'guizang-ppt-skill.zip' }
    if ($components -contains 'outlook') { $names += 'outlook-mcp.zip' }
    if ($components -contains 'excel') { $names += 'excel-mcp.zip' }
    return $names
}

function ConvertTo-ModelDisplayName {
    param([string]$ModelId, [string]$ProviderName = '')

    if ([string]::IsNullOrWhiteSpace($ModelId)) { return '' }
    $parts = @($ModelId -split '[-_/\.]+' | Where-Object { $_ })
    if ($parts.Count -eq 0) { return $ModelId }
    if (-not [string]::IsNullOrWhiteSpace($ProviderName)) {
        if ($parts[0] -ieq $ProviderName) { $parts[0] = $ProviderName }
        else { $parts = @($ProviderName) + $parts }
    }
    $display = @($parts | ForEach-Object {
        if ($_ -match '^v\d+$') { $_.ToUpperInvariant() }
        elseif ($_.Length -le 1) { $_.ToUpperInvariant() }
        else { $_.Substring(0, 1).ToUpperInvariant() + $_.Substring(1) }
    })
    return ($display -join ' ')
}

function Get-ApiModelsUrl {
    param([string]$BaseUrl)

    if ([string]::IsNullOrWhiteSpace($BaseUrl)) { return $null }
    return ($BaseUrl.TrimEnd('/') + '/models')
}

function ConvertFrom-ModelsApiResponse {
    param([string]$Json)

    $models = @()
    if ([string]::IsNullOrWhiteSpace($Json)) { return $models }
    try {
        $parsed = $Json | ConvertFrom-Json
    } catch {
        return $models
    }
    if ($null -eq $parsed) { return $models }

    $items = @()
    if ($parsed -is [array]) { $items = @($parsed) }
    elseif ($parsed -is [string]) { $items = @($parsed) }
    elseif ($parsed.PSObject.Properties['data']) { $items = @($parsed.data) }
    elseif ($parsed.PSObject.Properties['models']) { $items = @($parsed.models) }
    elseif ($parsed.PSObject.Properties['id']) { $items = @($parsed) }

    foreach ($item in $items) {
        $id = $null
        if ($item -is [string]) { $id = $item }
        elseif ($item -and $item.PSObject.Properties['id']) { $id = [string]$item.id }
        elseif ($item -and $item.PSObject.Properties['name']) { $id = [string]$item.name }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $id = $id.Trim()
        if ($models -notcontains $id) { $models += $id }
    }
    return $models
}

function Select-DefaultModel {
    param([string[]]$Models)

    $list = @($Models | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($list.Count -eq 0) { return $null }
    foreach ($model in $list) {
        if ($model -ieq 'deepseek-flash') { return $model }
    }
    foreach ($model in $list) {
        if ($model -like '*flash*') { return $model }
    }
    return $list[0]
}

function Get-ApiKeyReference {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    return '{file:' + ($Path -replace '\\', '/') + '}'
}

function Get-MajorVersion {
    param([string]$Version)

    if ([string]::IsNullOrWhiteSpace($Version)) { return $null }
    if ($Version -match '^\s*v?(\d+)') { return [int]$Matches[1] }
    return $null
}

function Get-OpenCodeLatestVersion {
    param(
        [string[]]$Registries = @('https://registry.npmmirror.com', 'https://registry.npmjs.org'),
        [int]$TimeoutSec = 5,
        [scriptblock]$Fetch
    )

    if (-not $Fetch) {
        $Fetch = {
            param($Url, $Timeout)
            Invoke-RestMethod -Uri $Url -TimeoutSec $Timeout -Headers @{ 'User-Agent' = 'opencode-installer' } -ErrorAction Stop
        }
    }

    foreach ($registry in @($Registries)) {
        if ([string]::IsNullOrWhiteSpace($registry)) { continue }
        $url = $registry.TrimEnd('/') + '/@opencode/cli/latest'
        try {
            $response = & $Fetch $url $TimeoutSec
            $text = if ($response -is [string]) { $response } elseif ($null -ne $response -and $response.PSObject.Properties['version']) { [string]$response.version } else { '' }
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $match = [regex]::Match($text, '\d+\.\d+\.\d+')
            if ($match.Success) { return $match.Value }
        } catch {
            continue
        }
    }
    return $null
}

function ConvertTo-MenuInput {
    param([string]$Text)

    if ($null -eq $Text) { return '' }
    $value = ([string]$Text) -replace '[\x00-\x1F\x7F]', ''
    $characters = foreach ($character in $value.ToCharArray()) {
        $code = [int][char]$character
        if ($code -ge 0xFF01 -and $code -le 0xFF5E) { [char]($code - 0xFEE0) } else { $character }
    }
    return ((-join $characters).Trim())
}

function Resolve-AiProviderChoice {
    param([string]$Choice)

    $normalized = ConvertTo-MenuInput -Text $Choice
    if ([string]::IsNullOrWhiteSpace($normalized)) { return 'deepseek' }
    switch ($normalized.ToLowerInvariant()) {
        '1' { return 'deepseek' }
        '2' { return 'custom' }
        'deepseek' { return 'deepseek' }
        'custom' { return 'custom' }
        default { return $null }
    }
}

function ConvertTo-CommandLineArgument {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '""' }
    if ($Value -match '[\s"]') { return '"' + ($Value -replace '"', '\"') + '"' }
    return $Value
}

function Invoke-CommandWithTimeout {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [int]$TimeoutSec = 300,
        [int]$HeartbeatSec = 15,
        [scriptblock]$OnHeartbeat
    )

    $argumentLine = (@($Arguments) | ForEach-Object { ConvertTo-CommandLineArgument ([string]$_) }) -join ' '
    if ($FilePath -match '\.(cmd|bat)$') {
        $executable = $env:ComSpec
        $processArguments = '/d /s /c "' + (ConvertTo-CommandLineArgument $FilePath) + $(if ($argumentLine) { ' ' + $argumentLine } else { '' }) + '"'
    } else {
        $executable = $FilePath
        $processArguments = $argumentLine
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $executable
    $startInfo.Arguments = $processArguments
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::Default
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::Default

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $outputTask = $null
    $errorTask = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $timedOut = $false
    try {
        if (-not $process.Start()) { throw "无法启动进程：$FilePath" }
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        $errorTask = $process.StandardError.ReadToEndAsync()
        $nextHeartbeat = if ($HeartbeatSec -gt 0) { $HeartbeatSec } else { [int]::MaxValue }
        while (-not $process.WaitForExit(1000)) {
            $elapsed = [int]$stopwatch.Elapsed.TotalSeconds
            if ($TimeoutSec -gt 0 -and $elapsed -ge $TimeoutSec) {
                $timedOut = $true
                try { & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null } catch { try { $process.Kill() } catch { } }
                try { $process.WaitForExit(5000) | Out-Null } catch { }
                break
            }
            if ($HeartbeatSec -gt 0 -and $elapsed -ge $nextHeartbeat) {
                if ($OnHeartbeat) { try { & $OnHeartbeat $elapsed } catch { } }
                $nextHeartbeat += $HeartbeatSec
            }
        }
        if (-not $process.HasExited) {
            try { $process.Kill() } catch { }
            try { $process.WaitForExit(5000) | Out-Null } catch { }
        }
    } finally {
        $stopwatch.Stop()
    }

    $output = if ($outputTask -and $outputTask.Wait(5000)) { [string]$outputTask.Result } else { '' }
    $errorOutput = if ($errorTask -and $errorTask.Wait(5000)) { [string]$errorTask.Result } else { '' }
    $exitCode = if ($timedOut) { -1 } elseif ($process.HasExited) { $process.ExitCode } else { -1 }
    try { $process.Dispose() } catch { }

    return [pscustomobject]@{
        ExitCode    = $exitCode
        TimedOut    = $timedOut
        Output      = $(if ($null -eq $output) { '' } else { $output.TrimEnd() })
        Error       = $(if ($null -eq $errorOutput) { '' } else { $errorOutput.TrimEnd() })
        DurationSec = [int]$stopwatch.Elapsed.TotalSeconds
    }
}

function Get-OpenCodeInstallMethod {
    param([string]$Path, [string]$NpmPrefix)

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($NpmPrefix)) { return 'unknown' }
    $normalizedPath = ($Path -replace '/', '\').TrimEnd('\')
    $normalizedPrefix = ($NpmPrefix -replace '/', '\').TrimEnd('\')
    if ($normalizedPath -ieq $normalizedPrefix -or $normalizedPath -like "$normalizedPrefix\*") { return 'npm' }
    return 'curl'
}

function Resolve-NpmExecutable {
    param([scriptblock]$Lookup)

    if (-not $Lookup) { $Lookup = { param($name) Get-Command $name -ErrorAction SilentlyContinue } }

    # Never resolve npm.ps1: it re-runs the original command line via
    # Invoke-Expression under strict mode, which breaks $script: arguments.
    foreach ($candidate in @('npm.cmd', 'npm.exe')) {
        $command = & $Lookup $candidate
        if (-not $command) { continue }
        if ($command -is [System.Management.Automation.CommandInfo]) {
            if ($command.CommandType -eq 'Application') { return [string]$command.Source }
            continue
        }
        return [string]$command
    }

    $fallback = & $Lookup 'npm'
    if (-not $fallback) { return $null }
    if ($fallback -is [System.Management.Automation.CommandInfo]) {
        if ($fallback.CommandType -eq 'Application') { return [string]$fallback.Source }
        return $null
    }
    return [string]$fallback
}

function Get-NpmCommand {
    param([scriptblock]$Lookup)

    $resolved = Resolve-NpmExecutable -Lookup $Lookup
    if ($resolved) { return $resolved }
    return 'npm'
}

function Test-SafeNpmCommand {
    param([string]$Command)

    if ([string]::IsNullOrWhiteSpace($Command)) { return $false }
    if ($Command -like '*.ps1') { return $false }
    try {
        $resolved = Get-Command $Command -ErrorAction Stop
        return [bool]($resolved -and $resolved.CommandType -eq 'Application')
    } catch {
        return $false
    }
}

function Get-ProxyConfiguration {
    param([string]$Uri = 'https://registry.npmmirror.com')

    foreach ($name in @('HTTPS_PROXY', 'https_proxy', 'HTTP_PROXY', 'http_proxy', 'ALL_PROXY', 'all_proxy')) {
        $value = [Environment]::GetEnvironmentVariable($name)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return [pscustomobject]@{ Uri = $value; Source = $name }
        }
    }

    try {
        $systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $target = [Uri]$Uri
        $proxy = $systemProxy.GetProxy($target)
        if ($proxy -and $proxy.AbsoluteUri -ne $target.AbsoluteUri -and $proxy.AbsoluteUri -ne '') {
            return [pscustomobject]@{ Uri = $proxy.AbsoluteUri; Source = 'system' }
        }
    } catch {
        $null = $_
    }
    return $null
}

function Invoke-InstallerWebRequest {
    param(
        [string]$Uri,
        [int]$TimeoutSec = 30,
        [string]$Method = 'GET',
        [string]$OutFile = ''
    )

    $proxy = Get-ProxyConfiguration -Uri $Uri
    $parameters = @{
        Uri             = $Uri
        UseBasicParsing = $true
        TimeoutSec      = $TimeoutSec
        ErrorAction     = 'Stop'
    }
    if ($Method -ne 'GET') { $parameters.Method = $Method }
    if ($OutFile) { $parameters.OutFile = $OutFile }
    if ($proxy) { $parameters.Proxy = $proxy.Uri }

    return Invoke-WebRequest @parameters
}

function Invoke-InstallerDownload {
    param(
        [string]$Uri,
        [string]$OutFile,
        [int]$TimeoutSec = 600
    )

    Invoke-InstallerWebRequest -Uri $Uri -OutFile $OutFile -TimeoutSec $TimeoutSec | Out-Null
}

function Test-UrlConnectivity {
    param(
        [string]$Uri,
        [int]$TimeoutSec = 8
    )

    $proxy = Get-ProxyConfiguration -Uri $Uri
    try {
        $parameters = @{
            Uri             = $Uri
            Method          = 'Head'
            UseBasicParsing = $true
            TimeoutSec      = $TimeoutSec
            ErrorAction     = 'Stop'
        }
        if ($proxy) { $parameters.Proxy = $proxy.Uri }
        $response = Invoke-WebRequest @parameters
        return [pscustomobject]@{
            Ok     = $true
            Status = [int]$response.StatusCode
            Error  = ''
            Proxy  = $proxy
        }
    } catch {
        return [pscustomobject]@{
            Ok     = $false
            Status = 0
            Error  = $_.Exception.Message
            Proxy  = $proxy
        }
    }
}

function Test-GitBashEnvironment {
    if (-not [string]::IsNullOrWhiteSpace($env:MSYSTEM)) { return $true }
    if ($env:OSTYPE -match '^(msys|cygwin)') { return $true }
    if (-not [string]::IsNullOrWhiteSpace($env:SHELL) -and $env:SHELL -match 'bash') { return $true }
    return $false
}

function Get-PowerShell7InstallDir {
    $base = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = $env:LOCALAPPDATA }
    return (Join-Path $base 'Programs\PowerShell\7')
}

function Get-PowerShell7DownloadUrls {
    $version = $script:PowerShell7Version
    $file = "PowerShell-$version-win-x64.zip"
    $urls = New-Object System.Collections.ArrayList

    if (-not [string]::IsNullOrWhiteSpace($env:OC_POWERSHELL_URL)) {
        [void]$urls.Add($env:OC_POWERSHELL_URL)
    }
    [void]$urls.Add("https://github.com/PowerShell/PowerShell/releases/download/v$version/$file")
    [void]$urls.Add("https://mirrors.tuna.tsinghua.edu.cn/github-release/PowerShell/PowerShell/v$version%20Release%20of%20PowerShell/$file")

    return @($urls)
}

function Get-PowerShell7LocalZipCandidates {
    param([string[]]$Directories)

    $file = "PowerShell-$($script:PowerShell7Version)-win-x64.zip"
    $found = @()
    foreach ($dir in @($Directories)) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        $path = Join-Path $dir $file
        if (Test-Path -LiteralPath $path) { $found += $path }
    }
    return $found
}

function Test-FileSha256 {
    param([string]$Path, [string]$Expected)

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Expected)) { return $false }
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
        return ($actual -ieq $Expected)
    } catch {
        return $false
    }
}

function Get-PowerShell7Version {
    param([string]$Path)

    $candidate = $Path
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $command = Get-Command pwsh.exe -ErrorAction SilentlyContinue
        if (-not $command) { $command = Get-Command pwsh -ErrorAction SilentlyContinue }
        if ($command) { $candidate = $command.Source }
        else { $candidate = Join-Path (Get-PowerShell7InstallDir) 'pwsh.exe' }
    }
    if ([string]::IsNullOrWhiteSpace($candidate) -or -not (Test-Path -LiteralPath $candidate)) { return $null }

    try {
        $raw = & $candidate -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>$null
        if ($LASTEXITCODE -eq 0 -and $raw) {
            $text = ([string]$raw).Trim()
            if ($text -match '^7\.\d+\.\d+') { return $text }
        }
    } catch {
        $null = $_
    }
    return $null
}

function Get-OpenCodeLauncherContent {
    return @'
@echo off
setlocal
set "TARGET=%~1"
if not defined TARGET set "TARGET=%CD%"
if exist "%TARGET%\*" goto :have_target
if exist "%TARGET%" for %%I in ("%TARGET%") do set "TARGET=%%~dpI"
:have_target
if "%TARGET:~-1%"=="\" set "TARGET=%TARGET:~0,-1%"
if not exist "%TARGET%\*" set "TARGET=%USERPROFILE%"
cd /d "%TARGET%" 2>nul || cd /d "%USERPROFILE%"
set "PWSH=%LOCALAPPDATA%\Programs\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" for %%I in (pwsh.exe) do set "PWSH=%%~$PATH:I"
if defined PWSH (
    "%PWSH%" -NoExit -Command opencode
) else (
    call opencode
    if errorlevel 1 pause
)
'@
}

function Get-OpenCodeIconPath {
    $candidates = @(
        (Join-Path $env:APPDATA 'npm\node_modules\@opencode\cli\bin\opencode.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\opencode\opencode.exe')
    )
    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) { return $path }
    }
    $command = Get-Command opencode -ErrorAction SilentlyContinue
    if ($command -and $command.Source -like '*.exe' -and (Test-Path -LiteralPath $command.Source)) {
        return $command.Source
    }
    return $null
}

function New-OpenCodeShortcut {
    param(
        [string]$DesktopPath = '',
        [string]$LauncherDir = ''
    )

    if ([string]::IsNullOrWhiteSpace($DesktopPath)) {
        $DesktopPath = [Environment]::GetFolderPath('Desktop')
    }
    if ([string]::IsNullOrWhiteSpace($DesktopPath)) {
        $DesktopPath = Join-Path $env:USERPROFILE 'Desktop'
    }
    if ([string]::IsNullOrWhiteSpace($LauncherDir)) {
        $LauncherDir = Join-Path $env:LOCALAPPDATA 'OpenCode'
    }

    try {
        New-Item -ItemType Directory -Force -Path $DesktopPath | Out-Null
        New-Item -ItemType Directory -Force -Path $LauncherDir | Out-Null

        $launcherPath = Join-Path $LauncherDir 'OpenCode.cmd'
        $launcherContent = (Get-OpenCodeLauncherContent) -replace "`r?`n", "`r`n"
        [System.IO.File]::WriteAllText($launcherPath, $launcherContent, [System.Text.Encoding]::ASCII)

        $shortcutPath = Join-Path $DesktopPath 'OpenCode.lnk'
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $launcherPath
        $shortcut.WorkingDirectory = $DesktopPath
        $shortcut.Description = '打开 OpenCode；可把文件夹拖到此快捷方式'
        $iconPath = Get-OpenCodeIconPath
        if ($iconPath) { $shortcut.IconLocation = "$iconPath,0" }
        $shortcut.Save()

        if (Test-Path -LiteralPath $shortcutPath) { return $shortcutPath }
    } catch {
        return $null
    }
    return $null
}

function Get-OpenCodeInstallPlan {
    param([hashtable]$State)

    $version = [string]$State.Version
    $path = [string]$State.Path
    $npmPrefix = [string]$State.NpmPrefix
    $hasV1 = [bool]$State.HasV1
    $hasV2 = [bool]$State.HasV2
    $latest = [string]$State.LatestVersion
    $major = Get-MajorVersion -Version $version
    $method = Get-OpenCodeInstallMethod -Path $path -NpmPrefix $npmPrefix
    $commands = New-Object System.Collections.ArrayList

    if (($null -ne $major -and $major -lt 2) -or $hasV1) {
        [void]$commands.Add([string[]]@('npm', 'uninstall', '-g', 'opencode-ai'))
        [void]$commands.Add([string[]]@('npm', 'install', '-g', '@opencode/cli@latest'))
        return @{
            Action   = 'MigrateV1'
            Reason   = '检测到 V1，先卸载 opencode-ai 再安装 @opencode/cli'
            Commands = @($commands)
        }
    }

    if ([string]::IsNullOrWhiteSpace($version) -and -not $hasV1) {
        [void]$commands.Add([string[]]@('npm', 'install', '-g', '@opencode/cli@latest'))
        $reason = if ($hasV2) { '检测到 @opencode/cli 但 opencode 命令不可用，重新安装以修复' } else { '未检测到 OpenCode，安装 @opencode/cli' }
        return @{
            Action   = 'Install'
            Reason   = $reason
            Commands = @($commands)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($version)) {
        if ([string]::IsNullOrWhiteSpace($latest)) {
            return @{
                Action   = 'None'
                Reason   = "已安装 v$version（无法获取最新版，保留现有版本）"
                Commands = @()
            }
        }

        $installedVersion = $null
        $latestVersion = $null
        $installedText = if ($version -match '(\d+\.\d+\.\d+)') { $Matches[1] } else { $version.TrimStart('v') }
        $latestText = if ($latest -match '(\d+\.\d+\.\d+)') { $Matches[1] } else { $latest.TrimStart('v') }
        if ([version]::TryParse($installedText, [ref]$installedVersion) -and [version]::TryParse($latestText, [ref]$latestVersion)) {
            if ($installedVersion -ge $latestVersion) {
                return @{ Action = 'None'; Reason = "已是最新版（v$version）"; Commands = @() }
            }
        } else {
            return @{
                Action   = 'None'
                Reason   = "无法比较版本（已安装 v$version，最新 v$latest），保留现有版本"
                Commands = @()
            }
        }
    }

    if ($method -eq 'npm') {
        [void]$commands.Add([string[]]@('npm', 'install', '-g', '@opencode/cli@latest'))
        return @{
            Action   = 'UpdateNpm'
            Reason   = "npm 安装，直接使用 npm 更新到 v$latest"
            Commands = @($commands)
        }
    }

    [void]$commands.Add([string[]]@('opencode', 'upgrade'))
    return @{
        Action   = 'UpdateCurl'
        Reason   = "系统安装，使用 opencode upgrade 更新到 v$latest"
        Commands = @($commands)
    }
}

function Get-ConfigSectionStyle {
    param([string]$ConfigJson, [string]$Section)

    $config = $null
    if (-not [string]::IsNullOrWhiteSpace($ConfigJson)) {
        try {
            $config = $ConfigJson | ConvertFrom-Json
        } catch {
            throw "opencode.json 内容不是合法 JSON：$($_.Exception.Message)"
        }
    }
    if ($null -eq $config) { return 'V2' }

    switch ($Section) {
        'provider' {
            if ($config.PSObject.Properties['providers']) { return 'V2' }
            if ($config.PSObject.Properties['provider']) { return 'V1' }
            return 'V2'
        }
        'mcp' {
            $mcpProperty = $config.PSObject.Properties['mcp']
            if (-not $mcpProperty -or $null -eq $mcpProperty.Value) { return 'V2' }
            $mcp = $mcpProperty.Value
            if ($mcp.PSObject.Properties['servers']) { return 'V2' }
            $directServers = @($mcp.PSObject.Properties | Where-Object { $_.Name -ne 'timeout' })
            if ($directServers.Count -gt 0) { return 'V1' }
            return 'V2'
        }
        'permissions' {
            if ($config.PSObject.Properties['permissions']) { return 'V2' }
            if ($config.PSObject.Properties['permission']) { return 'V1' }
            return 'V2'
        }
        default { throw "未知配置区段：$Section" }
    }
}

function Get-EffectivePermissionStyle {
    param([string]$ConfigJson)

    $style = Get-ConfigSectionStyle -ConfigJson $ConfigJson -Section 'permissions'
    if ($style -eq 'V2' -and $ConfigJson -notmatch '"permissions"\s*:') {
        $providerStyle = Get-ConfigSectionStyle -ConfigJson $ConfigJson -Section 'provider'
        if ($providerStyle -eq 'V1') { return 'V1' }
    }
    return $style
}

function New-ProviderPatch {
    param(
        [string]$Style = 'V2',
        [string]$ProviderId,
        [string]$Name,
        [string]$BaseUrl,
        [string]$ApiKeyRef = '',
        [string[]]$Models = @(),
        [string]$MainModel,
        [string]$TitleModel = ''
    )

    if ([string]::IsNullOrWhiteSpace($ProviderId)) { throw 'ProviderId 不能为空' }
    if ([string]::IsNullOrWhiteSpace($MainModel)) { throw 'MainModel 不能为空' }
    if ([string]::IsNullOrWhiteSpace($TitleModel)) { $TitleModel = $MainModel }

    $allModels = @()
    foreach ($model in (@($Models) + @($MainModel) + @($TitleModel))) {
        if (-not [string]::IsNullOrWhiteSpace($model) -and $allModels -notcontains $model) {
            $allModels += $model
        }
    }
    $modelMap = [ordered]@{}
    foreach ($model in $allModels) {
        $modelMap[$model] = [ordered]@{ name = (ConvertTo-ModelDisplayName -ModelId $model -ProviderName $Name) }
    }

    $settings = [ordered]@{ baseURL = $BaseUrl }
    if (-not [string]::IsNullOrWhiteSpace($ApiKeyRef)) { $settings.apiKey = $ApiKeyRef }

    if ($Style -eq 'V1') {
        $provider = [ordered]@{}
        $provider[$ProviderId] = [ordered]@{
            npm     = '@ai-sdk/openai-compatible'
            name    = $Name
            options = $settings
            models  = $modelMap
        }
        $patch = [ordered]@{
            model        = "$ProviderId/$MainModel"
            small_model  = "$ProviderId/$TitleModel"
            provider     = $provider
        }
    } else {
        $providers = [ordered]@{}
        $providers[$ProviderId] = [ordered]@{
            name     = $Name
            package  = '@opencode/ai/providers/openai-compatible'
            settings = $settings
            models   = $modelMap
        }
        $patch = [ordered]@{
            model     = "$ProviderId/$MainModel"
            providers = $providers
            agents    = [ordered]@{ title = [ordered]@{ model = "$ProviderId/$TitleModel" } }
        }
    }
    return ($patch | ConvertTo-Json -Depth 20)
}

function New-McpServerPatch {
    param(
        [string]$Style = 'V2',
        [string]$McpName,
        [string[]]$Command,
        [hashtable]$Environment = @{},
        [hashtable]$McpTimeout = $null
    )

    if ([string]::IsNullOrWhiteSpace($McpName)) { throw 'McpName 不能为空' }
    $server = [ordered]@{ type = 'local'; command = @($Command) }
    if ($Environment -and $Environment.Count -gt 0) { $server.environment = $Environment }

    if ($Style -eq 'V1') {
        $mcp = [ordered]@{}
        $mcp[$McpName] = $server
        return ([ordered]@{ mcp = $mcp } | ConvertTo-Json -Depth 20)
    }

    if ($null -eq $McpTimeout) { $McpTimeout = @{ startup = 60000; catalog = 60000 } }
    $servers = [ordered]@{}
    $servers[$McpName] = $server
    $mcp = [ordered]@{}
    if ($McpTimeout.Count -gt 0) { $mcp.timeout = $McpTimeout }
    $mcp.servers = $servers
    return ([ordered]@{ mcp = $mcp } | ConvertTo-Json -Depth 20)
}

function Add-SkillPermissionToJson {
    param([string]$ConfigJson, [string]$Style = 'V2')

    if ($Style -eq 'V1') {
        return Merge-OpenCodeConfigJson -ConfigJson $ConfigJson -PatchJson '{"permission":{"skill":{"*":"allow"}}}'
    }

    $config = $null
    if (-not [string]::IsNullOrWhiteSpace($ConfigJson)) {
        try {
            $config = $ConfigJson | ConvertFrom-Json
        } catch {
            throw "opencode.json 内容不是合法 JSON：$($_.Exception.Message)"
        }
    }
    if ($null -eq $config) { $config = New-Object PSObject }

    $permissionsProperty = $config.PSObject.Properties['permissions']
    $permissions = @()
    if ($permissionsProperty) {
        $permissions = @($permissionsProperty.Value)
    }

    $hasRule = $false
    foreach ($rule in $permissions) {
        if ($rule.action -eq 'skill' -and $rule.resource -eq '*' -and $rule.effect -eq 'allow') {
            $hasRule = $true
            break
        }
    }
    if (-not $hasRule) {
        $permissions += [pscustomobject]@{ action = 'skill'; resource = '*'; effect = 'allow' }
    }

    if ($permissionsProperty) { $config.permissions = $permissions }
    else { $config | Add-Member -MemberType NoteProperty -Name 'permissions' -Value $permissions }
    return ($config | ConvertTo-Json -Depth 100)
}

function Resolve-ConfigValueReference {
    param([string]$Value, [string]$ConfigDir)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    if ($Value -match '^\{file:(.+)\}$') {
        $path = $Matches[1]
        if ($path -match '^[~][/\\](.*)$') {
            $path = Join-Path $env:USERPROFILE $Matches[1]
        } elseif (-not [System.IO.Path]::IsPathRooted($path)) {
            $path = Join-Path $ConfigDir $path
        }
        if (-not (Test-Path -LiteralPath $path)) { return $null }
        try { return (Get-Content -LiteralPath $path -Raw -Encoding UTF8).Trim() }
        catch { return $null }
    }

    if ($Value -match '^\{env:([A-Za-z_][A-Za-z0-9_]*)\}$') {
        return [Environment]::GetEnvironmentVariable($Matches[1])
    }

    return $Value
}

function Get-DefaultComponentSelection {
    param([string]$Edition)

    $result = @()
    foreach ($key in @(Get-EnabledComponents -Edition $Edition)) {
        $result += @{ Key = $key; Selected = $true }
    }
    return $result
}

function Test-ComponentEnabled {
    param([string]$Edition, [string]$Key)

    return (@(Get-EnabledComponents -Edition $Edition) -contains $Key)
}

function Resolve-SkillAction {
    param([string]$InstalledVersion, [string]$PayloadVersion)

    if ([string]::IsNullOrWhiteSpace($InstalledVersion)) { return 'Install' }

    $installed = $null
    $payload = $null
    $isInstalledVersion = [version]::TryParse($InstalledVersion, [ref]$installed)
    $isPayloadVersion = [version]::TryParse($PayloadVersion, [ref]$payload)
    if (-not $isInstalledVersion -or -not $isPayloadVersion) { return 'Update' }
    if ($installed -lt $payload) { return 'Update' }
    return 'Skip'
}

function Read-InstallerManifest {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        return $null
    }
    if ($null -eq $json) { return $null }

    return @{
        installerVersion = [string]$json.installerVersion
        skills           = ConvertTo-PropertyHashtable $json.skills
    }
}

function Save-InstallerManifest {
    param([string]$Path, [hashtable]$Manifest)

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $Manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function ConvertTo-PropertyHashtable {
    param($InputObject)

    $result = @{}
    if ($null -eq $InputObject) { return $result }
    foreach ($property in $InputObject.PSObject.Properties) {
        $result[$property.Name] = $property.Value
    }
    return $result
}

function Get-SkillVersionFromSkillMd {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    foreach ($line in (Get-Content -LiteralPath $Path -TotalCount 40 -Encoding UTF8)) {
        if ($line -match '^\s*version:\s*"?([0-9][0-9\.]*)"?\s*$') { return $Matches[1] }
    }
    return $null
}

function Update-OpenCodeConfigJson {
    param([string]$ConfigJson, [string]$McpName, [string[]]$McpCommand, [bool]$AllowSkillPermission)

    $patch = @{ mcp = @{ $McpName = @{ type = 'local'; command = @($McpCommand) } } }
    if ($AllowSkillPermission) { $patch.permission = @{ skill = @{ '*' = 'allow' } } }
    return Merge-OpenCodeConfigJson -ConfigJson $ConfigJson -PatchJson ($patch | ConvertTo-Json -Depth 10)
}

function Add-SkillPermissionJson {
    param([string]$ConfigJson)

    return Merge-OpenCodeConfigJson -ConfigJson $ConfigJson -PatchJson '{"permission":{"skill":{"*":"allow"}}}'
}

function Merge-OpenCodeConfigJson {
    param([string]$ConfigJson, [string]$PatchJson)

    $config = $null
    if (-not [string]::IsNullOrWhiteSpace($ConfigJson)) {
        try {
            $config = $ConfigJson | ConvertFrom-Json
        } catch {
            throw "opencode.json 内容不是合法 JSON：$($_.Exception.Message)"
        }
    }
    if ($null -eq $config) { $config = New-Object PSObject }

    if (-not [string]::IsNullOrWhiteSpace($PatchJson)) {
        $patch = $null
        try {
            $patch = $PatchJson | ConvertFrom-Json
        } catch {
            throw "补丁内容不是合法 JSON：$($_.Exception.Message)"
        }
        if ($null -ne $patch) { Merge-ObjectDeep -Base $config -Patch $patch | Out-Null }
    }
    return ($config | ConvertTo-Json -Depth 100)
}

function Merge-ObjectDeep {
    param($Base, $Patch)

    foreach ($property in $Patch.PSObject.Properties) {
        $existing = $Base.PSObject.Properties[$property.Name]
        $patchIsObject = $null -ne $property.Value -and $property.Value.GetType().Name -eq 'PSCustomObject'
        $baseIsObject = $null -ne $existing -and $existing.Value -and $existing.Value.GetType().Name -eq 'PSCustomObject'
        if ($patchIsObject -and $baseIsObject) {
            Merge-ObjectDeep -Base $existing.Value -Patch $property.Value | Out-Null
        } elseif ($null -ne $existing) {
            $Base.PSObject.Properties[$property.Name].Value = $property.Value
        } else {
            $Base | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }
    return $Base
}

function Get-NodeDownloadCandidates {
    param([string]$IndexJson, [string]$BaseUrl = 'https://npmmirror.com/mirrors/node')

    $fallbacks = @(
        "$BaseUrl/v22.19.0/node-v22.19.0-x64.msi",
        "$BaseUrl/v20.19.0/node-v20.19.0-x64.msi",
        'https://nodejs.org/dist/v22.19.0/node-v22.19.0-x64.msi'
    )

    $dynamic = @()
    try {
        $entries = $IndexJson | ConvertFrom-Json
        foreach ($entry in @($entries)) {
            $rawVersion = [string]$entry.version
            if ([string]::IsNullOrWhiteSpace($rawVersion)) { continue }
            if ($null -eq $entry.lts -or $entry.lts -eq $false) { continue }
            $parsed = $null
            if (-not [version]::TryParse($rawVersion.TrimStart('v'), [ref]$parsed)) { continue }
            if ($dynamic.Count -eq 0 -or $parsed -gt $dynamic[0].Parsed) {
                $dynamic = @(@{ Parsed = $parsed; Url = "$BaseUrl/$rawVersion/node-$rawVersion-x64.msi" })
            }
        }
    } catch { $null = $_ }

    $candidates = @()
    if ($dynamic.Count -gt 0) { $candidates += $dynamic[0].Url }
    return @($candidates + $fallbacks)
}

function Get-PayloadRoot {
    param([string]$TempRoot, [string]$AssetsDir, [string[]]$RequiredNames)

    if ($TempRoot -and (Test-Path -LiteralPath $TempRoot)) {
        $missing = @(Get-MissingPayloads -Root $TempRoot -Names $RequiredNames)
        if ($missing.Count -eq 0) { return $TempRoot }
    }
    return $AssetsDir
}

function Get-HostPortFromUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    $uri = $null
    if (-not [uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) { return $null }
    return @{ Host = $uri.Host; Port = $uri.Port }
}

function Test-TcpEndpoint {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync($HostName, $Port)
        if (-not $task.Wait($TimeoutMs)) { return $false }
        return $client.Connected
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Resolve-ApiFailureMessage {
    param([int]$StatusCode = 0, [string]$Endpoint = '')

    switch ($StatusCode) {
        401 { return 'API Key 无效（401），请检查是否复制完整' }
        403 { return '该 Key 无权限或账户欠费（403）' }
        429 { return '请求过于频繁（429），请稍后再试' }
        0 { return "无法连接 $Endpoint（可能是公司网络/代理拦截，或本机断网）" }
        default { return "API 返回错误（$StatusCode）" }
    }
}

function Expand-SkillPack {
    param([string]$ZipPath, [string]$InnerRoot, [string]$Destination)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $staging = Join-Path ([System.IO.Path]::GetTempPath()) ("opencode-skill-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $staging | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $staging)
        $inner = Join-Path $staging ($InnerRoot -replace '/', '\')
        if (-not (Test-Path -LiteralPath $inner)) { throw "技能包内未找到目录：$InnerRoot" }

        $preservedEnv = $null
        $envFile = Join-Path $Destination '.env'
        if (Test-Path -LiteralPath $envFile) { $preservedEnv = [System.IO.File]::ReadAllText($envFile) }

        if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Recurse -Force }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Destination) | Out-Null
        Move-Item -LiteralPath $inner -Destination $Destination

        if ($null -ne $preservedEnv) {
            [System.IO.File]::WriteAllText((Join-Path $Destination '.env'), $preservedEnv, (New-Object System.Text.UTF8Encoding($false)))
        }
    } finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Expand-ZipPackage {
    param([string]$ZipPath, [string]$Destination)

    if (-not (Test-Path -LiteralPath $ZipPath)) { throw "压缩包不存在：$ZipPath" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Recurse -Force }
    $parent = Split-Path -Parent $Destination
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
}

function Expand-McpPackage {
    param([string]$ZipPath, [string]$PackageRoot)

    if (Test-Path -LiteralPath $ZipPath) {
        Expand-ZipPackage -ZipPath $ZipPath -Destination $PackageRoot
        return $true
    }
    return (Test-Path -LiteralPath $PackageRoot)
}

function Get-McpVenvDir {
    param([string]$McpRoot, [string]$Key)

    if ([string]::IsNullOrWhiteSpace($McpRoot)) { throw 'McpRoot 不能为空' }
    if ([string]::IsNullOrWhiteSpace($Key)) { throw 'Key 不能为空' }
    return (Join-Path $McpRoot ("venv-" + $Key))
}

function Get-PackageVersionFromPyProject {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $inProject = $false
        foreach ($line in (Get-Content -LiteralPath $Path -Encoding UTF8)) {
            if ($line -match '^\s*\[project\]\s*$') { $inProject = $true; continue }
            if ($line -match '^\s*\[') { $inProject = $false; continue }
            if ($inProject -and $line -match '^\s*version\s*=\s*["'']([^"'']+)["'']') { return $Matches[1] }
        }
    } catch {
        return $null
    }
    return $null
}

function Get-PipInstallArguments {
    param([string]$RequirementsPath, [string]$IndexUrl)

    $pipArguments = @('-m', 'pip', 'install', '-r', $RequirementsPath)
    if (-not [string]::IsNullOrWhiteSpace($IndexUrl)) { $pipArguments += @('-i', $IndexUrl) }
    return $pipArguments
}

function Get-MissingPayloads {
    param([string]$Root, [string[]]$Names)

    $missing = @()
    foreach ($name in $Names) {
        if (-not (Test-Path -LiteralPath (Join-Path $Root $name))) { $missing += $name }
    }
    return $missing
}
