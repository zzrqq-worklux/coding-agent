# ============================================================
# OpenCode 一键安装器 v4.2.2
# 双版本：便携版（仅 MCP）/ 厚重版（PPT 双技能 + MCP）
# 适配 OpenCode V2、DeepSeek 动态模型发现
# ============================================================
[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$SelfTest,
    [string]$TestRoot = '',
    [string]$PayloadRoot = '',
    [switch]$NoPause
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { $null = $_ }
try { $OutputEncoding = [System.Text.Encoding]::UTF8 } catch { $null = $_ }
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { $null = $_ }
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch { $null = $_ }
$ErrorActionPreference = 'Continue'

$INSTALLER_VERSION = '4.2.2'
$PAYLOAD_DIR = Join-Path $env:TEMP 'opencode-installer-v4'
$NPM_MIRROR = 'https://registry.npmmirror.com'
$PIP_MIRROR = 'https://pypi.tuna.tsinghua.edu.cn/simple'
$DEEPSEEK_BASE_URL = 'https://api.deepseek.com/v1'

$script:DryRun = [bool]$DryRun
$script:SelfTest = [bool]$SelfTest
$script:IsAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$script:IsUserMode = -not $script:IsAdmin
$script:NoPause = $NoPause -or $DryRun -or $SelfTest
$script:InstallerVersion = $INSTALLER_VERSION
$script:NpmMirror = $NPM_MIRROR
$script:PipMirror = $PIP_MIRROR
$script:ConfigJson = ''
$script:ComponentResults = @()

# ------------------------------------------------------------
# 加载核心与 UI
# ------------------------------------------------------------
$coreCandidates = @()
if ($PSScriptRoot) { $coreCandidates += (Join-Path $PSScriptRoot 'lib\installer-core.ps1') }
if (Test-Path -LiteralPath $PAYLOAD_DIR) { $coreCandidates += (Join-Path $PAYLOAD_DIR 'installer-core.ps1') }
$corePath = ''
foreach ($candidate in $coreCandidates) {
    if (Test-Path -LiteralPath $candidate) { $corePath = $candidate; break }
}
if (-not $corePath) {
    Write-Host '  安装器组件缺失，请重新下载安装器' -ForegroundColor Red
    exit 1
}
. $corePath
$script:CorePath = $corePath

if (-not $script:InstallerCoreVersion -or $script:InstallerCoreVersion -ne $INSTALLER_VERSION) {
    Write-Host "  安装器核心组件版本不匹配（核心：$($script:InstallerCoreVersion)，安装器：$INSTALLER_VERSION）" -ForegroundColor Red
    Write-Host "  已加载：$corePath" -ForegroundColor DarkGray
    Write-Host '  请删除临时目录后重新运行，或重新下载安装器。' -ForegroundColor Yellow
    exit 1
}

$script:NpmExe = Resolve-NpmExecutable
if (-not $script:NpmExe) { $script:NpmExe = Get-NpmCommand }

$editionCandidates = @()
if (Test-Path -LiteralPath $PAYLOAD_DIR) { $editionCandidates += (Join-Path $PAYLOAD_DIR 'edition.txt') }
if ($PSScriptRoot) { $editionCandidates += (Join-Path $PSScriptRoot 'edition.txt') }
$INSTALLER_EDITION = Get-EditionFromFile -Paths $editionCandidates
$editionInfo = Get-EditionInfo -Edition $INSTALLER_EDITION
$script:EnabledComponents = @(Get-EnabledComponents -Edition $INSTALLER_EDITION)
$requiredPayloads = @(Get-RequiredPayloads -Edition $INSTALLER_EDITION)

if (-not $PayloadRoot) {
    $assetsDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'assets' } else { '' }
    $PayloadRoot = Get-PayloadRoot -TempRoot $PAYLOAD_DIR -AssetsDir $assetsDir -RequiredNames $requiredPayloads
}
$script:PayloadRoot = $PayloadRoot

$uiCandidates = @()
if ($PSScriptRoot) { $uiCandidates += (Join-Path $PSScriptRoot 'lib\ui.ps1') }
if (Test-Path -LiteralPath $PAYLOAD_DIR) { $uiCandidates += (Join-Path $PAYLOAD_DIR 'ui.ps1') }
$uiPath = ''
foreach ($candidate in $uiCandidates) {
    if (Test-Path -LiteralPath $candidate) { $uiPath = $candidate; break }
}
if (-not $uiPath) {
    Write-Host '  安装器界面组件缺失，请重新下载安装器' -ForegroundColor Red
    exit 1
}
. $uiPath

$script:ConfigRoot = if ($TestRoot) { $TestRoot } else { Join-Path $env:USERPROFILE '.config\opencode' }
$script:SkillsDir = Join-Path $script:ConfigRoot 'skills'
$script:ConfigFile = Join-Path $script:ConfigRoot 'opencode.json'
$script:ManifestPath = Join-Path $script:ConfigRoot 'installer-manifest.json'
$script:McpDir = Join-Path $script:ConfigRoot 'mcp-servers'
# 每个 MCP 使用独立 venv（venv-excel / venv-outlook），避免 mcp / fastmcp 版本互相破坏
$script:SecretsDir = Join-Path $script:ConfigRoot '.secrets'
$script:LogFile = Join-Path $env:TEMP ("opencode-installer-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + ".log")

$script:PptMaster = @{
    Key       = 'ppt-master'
    Name      = 'ppt-master'
    Display   = 'PPT-Master（原生 PPTX，PowerPoint 里可继续编辑）'
    Zip       = 'ppt-master-skill-v6.7.3.zip'
    InnerRoot = 'ppt-master/skills/ppt-master'
    Version   = '6.7.3'
    NeedsPip  = $true
    Credit    = 'PPT-Master by Hugo He · MIT · github.com/hugohe3/ppt-master'
}
$script:GuiZang = @{
    Key       = 'guizang'
    Name      = 'guizang-ppt-skill'
    Display   = '归藏网页 PPT（单文件 HTML 幻灯片，浏览器直接放映）'
    Zip       = 'guizang-ppt-skill.zip'
    InnerRoot = 'guizang-ppt-skill'
    Version   = '2026.09.20'
    NeedsPip  = $false
    Credit    = '归藏网页 PPT by 歸藏 · AGPL-3.0 · github.com/op7418/guizang-ppt-skill'
}
$script:OutlookMcp = @{
    Key         = 'outlook'
    Name        = 'outlook'
    Display     = 'Outlook 邮件助手（本机 Outlook COM，免云端认证）'
    Zip         = 'outlook-mcp.zip'
    PackageDir  = 'outlook-mcp'
    ProbeModule = 'outlook_com_mcp.server'
    Version     = '0.1.0'
}
$script:ExcelMcp = @{
    Key         = 'excel'
    Name        = 'my-excel'
    Display     = 'Excel 表格助手（读写 / 计算 / 图表 / 数据清洗）'
    Zip         = 'excel-mcp.zip'
    PackageDir  = 'excel-mcp'
    ProbeModule = 'excel_mcp.server'
    Version     = '0.1.8.post1'
}
$script:ComponentCatalog = @{
    'ppt-master' = $script:PptMaster.Display
    'guizang'    = $script:GuiZang.Display
    'outlook'    = $script:OutlookMcp.Display
    'excel'      = $script:ExcelMcp.Display
}

# ============================================================
# 基础环境
# ============================================================

function Get-OsCaption {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os.Caption) { return [string]$os.Caption }
    } catch { $null = $_ }
    return [Environment]::OSVersion.VersionString
}

function Get-NodeVersion {
    try {
        $version = & node --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) { return ([string]$version).Trim() }
    } catch { $null = $_ }
    return $null
}

function Get-PythonVersion {
    try {
        $version = & python --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) { return ([string]$version).Trim() }
    } catch { $null = $_ }
    return $null
}

function Get-OpenCodeVersion {
    try {
        $version = & opencode --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) {
            $text = ([string]$version).Trim()
            if ($text -match '(\d+\.\d+\.\d+)') { return $Matches[1] }
            return $text
        }
    } catch { $null = $_ }
    return $null
}

function Get-OpenCodePath {
    try {
        $command = Get-Command opencode -ErrorAction SilentlyContinue
        if ($command -and $command.Source) { return [string]$command.Source }
        if ($command -and $command.Path) { return [string]$command.Path }
    } catch { $null = $_ }
    return ''
}

function Show-DownloadDiagnostics {
    param([string[]]$Urls)

    if (-not $Urls -or $Urls.Count -eq 0) { return }
    $proxy = Get-ProxyConfiguration -Uri $Urls[0]
    if ($proxy) { Write-Status 'info' "检测到网络代理：$($proxy.Source) $($proxy.Uri)" }
    foreach ($url in $Urls) {
        $result = Test-UrlConnectivity -Uri $url -TimeoutSec 8
        if ($result.Ok) {
            Write-Status 'ok' "可访问：$url（HTTP $($result.Status)）"
        } else {
            Write-Status 'warn' "不可访问：$url（$($result.Error)）"
        }
    }
}

function Test-InstallerEnvironment {
    $checks = New-Object System.Collections.ArrayList

    [void]$checks.Add([pscustomobject]@{
        Name   = 'PowerShell'
        Ok     = $true
        Detail = "$($PSVersionTable.PSVersion) / $($PSVersionTable.PSEdition) / $($Host.Name)"
    })

    $gitBash = Test-GitBashEnvironment
    [void]$checks.Add([pscustomobject]@{
        Name   = 'Git Bash'
        Ok     = $true
        Detail = $(if ($gitBash) { '是（跳过 PowerShell 7 安装）' } else { '否' })
    })

    $pwshVersion = Get-PowerShell7Version
    [void]$checks.Add([pscustomobject]@{
        Name   = 'PowerShell 7'
        Ok     = $true
        Detail = $(if ($pwshVersion) { $pwshVersion } elseif ($gitBash) { '跳过（Git Bash）' } else { '未安装（可选）' })
    })

    $coreOk = ($script:InstallerCoreVersion -eq $INSTALLER_VERSION)
    [void]$checks.Add([pscustomobject]@{
        Name   = '核心组件'
        Ok     = $coreOk
        Detail = "$($script:InstallerCoreVersion)（期望 $INSTALLER_VERSION）@ $($script:CorePath)"
    })

    $required = @(
        'Resolve-NpmExecutable', 'Test-SafeNpmCommand', 'Get-OpenCodeInstallPlan', 'Get-ProxyConfiguration', 'Test-UrlConnectivity',
        'Test-GitBashEnvironment', 'Get-PowerShell7Version', 'Get-PowerShell7DownloadUrls', 'Get-PowerShell7LocalZipCandidates', 'Test-FileSha256',
        'Get-OpenCodeLauncherContent', 'New-OpenCodeShortcut', 'Get-McpVenvDir', 'Expand-McpPackage'
    )
    $missing = @($required | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    [void]$checks.Add([pscustomobject]@{
        Name   = '核心函数'
        Ok     = ($missing.Count -eq 0)
        Detail = $(if ($missing.Count -eq 0) { '齐全' } else { '缺少 ' + ($missing -join ', ') })
    })

    $npmCommand = Get-Command npm -ErrorAction SilentlyContinue
    $npmExists = [bool]$npmCommand
    $npmSafe = $npmExists -and (Test-SafeNpmCommand -Command $script:NpmExe)
    [void]$checks.Add([pscustomobject]@{
        Name   = 'npm 命令'
        Ok     = ((-not $npmExists) -or $npmSafe)
        Detail = $(if (-not $npmExists) { '未安装（跳过 npm 检查）' }
                   elseif ($npmSafe) { "$($script:NpmExe)（安全）" }
                   else { "不安全：$($script:NpmExe)（npm.ps1 会导致变量作用域错误）" })
    })

    $nodeVersion = Get-NodeVersion
    [void]$checks.Add([pscustomobject]@{
        Name   = 'Node.js'
        Ok     = $true
        Detail = $(if ($nodeVersion) { $nodeVersion } else { '未安装（安装器会尝试安装）' })
    })

    $pythonVersion = Get-PythonVersion
    [void]$checks.Add([pscustomobject]@{
        Name   = 'Python'
        Ok     = $true
        Detail = $(if ($pythonVersion) { $pythonVersion } else { '未安装（可选）' })
    })

    $openCodeVersion = Get-OpenCodeVersion
    [void]$checks.Add([pscustomobject]@{
        Name   = 'OpenCode'
        Ok     = $true
        Detail = $(if ($openCodeVersion) { $openCodeVersion } else { '未安装（安装器会安装）' })
    })

    $proxy = Get-ProxyConfiguration
    [void]$checks.Add([pscustomobject]@{
        Name   = '网络代理'
        Ok     = $true
        Detail = $(if ($proxy) { "$($proxy.Source) $($proxy.Uri)" } else { '未检测到代理（直连）' })
    })

    return $checks
}

function Invoke-SelfTest {
    $checks = @(Test-InstallerEnvironment)
    foreach ($check in $checks) {
        $icon = if ($check.Ok) { '[✓]' } else { '[×]' }
        $color = if ($check.Ok) { 'Green' } else { 'Red' }
        Write-Host ("  {0} {1}：{2}" -f $icon, $check.Name, $check.Detail) -ForegroundColor $color
    }
    return (@($checks | Where-Object { -not $_.Ok }).Count -eq 0)
}

function Get-NpmPrefix {
    try {
        $result = Invoke-CommandWithTimeout -FilePath $script:NpmExe -Arguments @('config', 'get', 'prefix') -TimeoutSec 15 -HeartbeatSec 0
        if (-not $result.TimedOut -and $result.ExitCode -eq 0 -and $result.Output) { return $result.Output.Trim() }
    } catch { $null = $_ }
    return (Join-Path $env:APPDATA 'npm')
}

function Get-NpmGlobalPackages {
    try {
        $result = Invoke-CommandWithTimeout -FilePath $script:NpmExe -Arguments @('ls', '-g', '--depth=0', '--json') -TimeoutSec 30 -HeartbeatSec 0
        if ($result.TimedOut) { return @() }
        $raw = [string]$result.Output
        $start = $raw.IndexOf('{')
        if ($start -lt 0) { return @() }
        $parsed = $raw.Substring($start) | ConvertFrom-Json
        if (-not $parsed.dependencies) { return @() }
        return @($parsed.dependencies.PSObject.Properties.Name)
    } catch { $null = $_ }
    return @()
}

function Update-ProcessPath {
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    $machinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $env:Path = "$machinePath;$userPath;$env:Path"
}

function Add-UserPathEntry {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    if ($userPath -notlike "*$Path*") {
        [Environment]::SetEnvironmentVariable('PATH', "$userPath;$Path", 'User')
    }
    if ($env:Path -notlike "*$Path*") { $env:Path = "$env:Path;$Path" }
}

function Test-WingetAvailable {
    try { $null = Get-Command winget -ErrorAction Stop; return $true } catch { return $false }
}

function Restart-Self {
    if ($script:DryRun) { return }
    $exePath = ''
    try { $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { $null = $_ }
    if ($exePath -and $exePath.EndsWith('.exe')) {
        Write-Status 'info' '正在自动重启安装器以刷新环境...'
        Start-Sleep 1
        Start-Process $exePath
    } else {
        Write-Status 'info' '请关闭本窗口后重新运行安装器'
    }
    exit 0
}

# ============================================================
# Node.js / Python
# ============================================================

function Install-NodeJS {
    Write-Status 'info' 'Node.js 未检测到，准备安装...'
    if ($script:DryRun) { Write-Status 'plan' '安装 Node.js LTS（winget 优先，失败则用国内镜像 MSI）'; return $true }

    if (Test-WingetAvailable) {
        Write-Status 'info' '通过 winget 安装 Node.js LTS...'
        $scope = if ($script:IsUserMode) { '--scope user ' } else { '' }
        $proc = Start-Process winget -ArgumentList "install OpenJS.NodeJS.LTS $scope--accept-source-agreements --accept-package-agreements --silent" -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            Update-ProcessPath
            $version = Get-NodeVersion
            if ($version) { Write-Status 'ok' "Node.js 安装完成（$version）"; return $true }
        }
        Write-Status 'warn' 'winget 方式未完成，改用镜像下载安装包...'
    }

    Show-DownloadDiagnostics -Urls @('https://npmmirror.com/mirrors/node/index.json', 'https://nodejs.org/dist/index.json')

    $indexJson = ''
    try {
        $indexJson = (Invoke-InstallerWebRequest -Uri 'https://npmmirror.com/mirrors/node/index.json' -TimeoutSec 30).Content
    } catch {
        Write-Status 'warn' "Node.js 版本索引下载失败：$($_.Exception.Message)"
    }
    $candidates = @(Get-NodeDownloadCandidates -IndexJson $indexJson)
    $msiPath = Join-Path $PAYLOAD_DIR 'nodejs-lts-x64.msi'
    $downloaded = $false
    foreach ($url in $candidates) {
        try {
            Write-Status 'info' "正在下载 $($url.Split('/')[-1])..."
            Invoke-InstallerDownload -Uri $url -OutFile $msiPath -TimeoutSec 600
            $downloaded = $true
            break
        } catch {
            Write-Status 'warn' "下载失败：$url（$($_.Exception.Message)）"
        }
    }
    if (-not $downloaded) {
        Write-Status 'error' 'Node.js 下载失败，请检查网络/代理后重新运行'
        Write-Status 'info' '如公司网络需要代理，请先设置 HTTPS_PROXY 环境变量；也可手动安装 https://nodejs.org/zh-cn/'
        return $false
    }

    $msiArgs = if ($script:IsUserMode) { "/i `"$msiPath`" /quiet /norestart MSIINSTALLPERUSER=1" } else { "/i `"$msiPath`" /quiet /norestart" }
    $proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow
    Remove-Item -LiteralPath $msiPath -Force -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) { Write-Status 'warn' "MSI 安装返回退出码 $($proc.ExitCode)" }
    Update-ProcessPath
    $version = Get-NodeVersion
    if ($version) { Write-Status 'ok' "Node.js 安装完成（$version）"; return $true }

    Write-Status 'info' 'Node.js 已安装，正在重新启动安装器以刷新环境...'
    Restart-Self
    return $false
}

function Install-Python {
    Write-Status 'info' 'Python 未检测到，准备安装...'
    if ($script:DryRun) { Write-Status 'plan' '安装 Python 3.12（winget 优先，失败则用国内镜像安装包）'; return $true }

    if (Test-WingetAvailable) {
        Write-Status 'info' '通过 winget 安装 Python 3.12...'
        $proc = Start-Process winget -ArgumentList 'install Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent' -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            Update-ProcessPath
            $version = Get-PythonVersion
            if ($version) { Write-Status 'ok' "Python 安装完成（$version）"; return $true }
        }
        Write-Status 'warn' 'winget 方式未完成，改用镜像下载安装包...'
    }

    $urls = @(
        'https://npmmirror.com/mirrors/python/3.12.9/python-3.12.9-amd64.exe',
        'https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe'
    )
    Show-DownloadDiagnostics -Urls $urls
    $installerPath = Join-Path $PAYLOAD_DIR 'python-installer.exe'
    $downloaded = $false
    foreach ($url in $urls) {
        try {
            Write-Status 'info' '正在下载 Python...'
            Invoke-InstallerDownload -Uri $url -OutFile $installerPath -TimeoutSec 600
            $downloaded = $true
            break
        } catch {
            Write-Status 'warn' "下载失败：$url（$($_.Exception.Message)）"
        }
    }
    if (-not $downloaded) {
        Write-Status 'error' 'Python 下载失败，请检查网络/代理后重新运行'
        Write-Status 'info' '如公司网络需要代理，请先设置 HTTPS_PROXY 环境变量；也可手动安装 https://www.python.org/'
        return $false
    }

    $scope = if ($script:IsUserMode) { 'InstallAllUsers=0' } else { 'InstallAllUsers=1' }
    $proc = Start-Process $installerPath -ArgumentList "/quiet $scope PrependPath=1 Include_pip=1" -Wait -PassThru -NoNewWindow
    Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
    if ($proc.ExitCode -ne 0) { Write-Status 'warn' "Python 安装返回退出码 $($proc.ExitCode)" }
    Update-ProcessPath
    $version = Get-PythonVersion
    if ($version) { Write-Status 'ok' "Python 安装完成（$version）"; return $true }

    Write-Status 'info' 'Python 已安装，正在重新启动安装器以刷新环境...'
    Restart-Self
    return $false
}

function Install-PowerShell7 {
    $existing = Get-PowerShell7Version
    if ($existing) {
        $existingDir = Get-PowerShell7InstallDir
        if (Test-Path -LiteralPath (Join-Path $existingDir 'pwsh.exe')) { Add-UserPathEntry -Path $existingDir }
        Write-Status 'ok' "PowerShell 7 已安装（$existing）"
        return $true
    }
    if (Test-GitBashEnvironment) {
        Write-Status 'skip' '检测到 Git Bash，跳过 PowerShell 7 安装'
        return $true
    }
    if ($script:DryRun) {
        Write-Status 'plan' "PowerShell 7：$($script:PowerShell7Version) 用户目录安装（无需管理员）"
        return $true
    }
    if (-not (Confirm-YesNo '是否安装 PowerShell 7？（可选，改善终端显示）' $true)) {
        Write-Status 'skip' '已跳过 PowerShell 7'
        return $true
    }

    if (-not (Test-Path -LiteralPath $PAYLOAD_DIR)) { New-Item -ItemType Directory -Force -Path $PAYLOAD_DIR | Out-Null }

    $installerDir = ''
    try { $installerDir = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch { $null = $_ }
    $localDirs = @($installerDir, $PSScriptRoot, $PAYLOAD_DIR) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $zipPath = Join-Path $PAYLOAD_DIR ("PowerShell-$($script:PowerShell7Version)-win-x64.zip")
    $verified = $false

    foreach ($candidate in @(Get-PowerShell7LocalZipCandidates -Directories $localDirs)) {
        if (Test-FileSha256 -Path $candidate -Expected $script:PowerShell7Sha256) {
            Write-Status 'ok' "使用本地 PowerShell 7 安装包：$candidate"
            if ($candidate -ne $zipPath) { Copy-Item -LiteralPath $candidate -Destination $zipPath -Force }
            $verified = $true
            break
        }
        Write-Status 'warn' "本地安装包校验失败，已忽略：$candidate"
    }

    if (-not $verified) {
        foreach ($url in @(Get-PowerShell7DownloadUrls)) {
            try {
                Write-Status 'info' "正在下载 PowerShell 7：$url"
                if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
                Invoke-InstallerDownload -Uri $url -OutFile $zipPath -TimeoutSec 600
                if (Test-FileSha256 -Path $zipPath -Expected $script:PowerShell7Sha256) {
                    $verified = $true
                    break
                }
                Write-Status 'warn' "SHA256 校验失败，尝试下一个来源：$url"
                Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Status 'warn' "下载失败：$url（$($_.Exception.Message)）"
            }
        }
    }

    if (-not $verified) {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
        Write-Status 'warn' 'PowerShell 7 安装包下载失败，跳过（不影响 OpenCode 安装）'
        Write-Status 'info' "可手动下载 PowerShell-$($script:PowerShell7Version)-win-x64.zip，放到安装器同目录后重跑。"
        return $false
    }

    $targetDir = Get-PowerShell7InstallDir
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path -LiteralPath $targetDir) { Remove-Item -LiteralPath $targetDir -Recurse -Force }
        New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $targetDir)
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Status 'warn' "PowerShell 7 解压失败：$($_.Exception.Message)"
        return $false
    }

    $pwshPath = Join-Path $targetDir 'pwsh.exe'
    $installedVersion = Get-PowerShell7Version -Path $pwshPath
    if (-not $installedVersion) {
        Write-Status 'warn' "PowerShell 7 已解压，但无法运行：$pwshPath"
        return $false
    }

    Add-UserPathEntry -Path $targetDir

    Write-Status 'ok' "PowerShell 7 已安装（$installedVersion）→ $targetDir"
    return $true
}

# ============================================================
# 镜像与 OpenCode V2
# ============================================================

function Set-NpmMirror {
    if (-not (Test-SafeNpmCommand -Command $script:NpmExe)) {
        Write-Status 'warn' "npm 命令不安全，跳过镜像设置：$($script:NpmExe)"
        return
    }
    if ($script:DryRun) { Write-Status 'plan' "npm config set registry $($script:NpmMirror)"; return }
    & $script:NpmExe config set registry $script:NpmMirror | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Status 'ok' 'npm 镜像已设置为 npmmirror.com'
    } else {
        Write-Status 'warn' "npm 镜像设置失败（退出码 $LASTEXITCODE）"
    }
}

function Set-PipMirror {
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) { return }
    if ($script:DryRun) { Write-Status 'plan' "python -m pip config set global.index-url $($script:PipMirror)"; return }
    & python -m pip config set global.index-url $script:PipMirror 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Status 'ok' 'pip 镜像已设置为清华源' }
    else { Write-Status 'warn' "pip 镜像设置失败（退出码 $LASTEXITCODE）" }
}

function Set-NpmUserPrefix {
    if (-not $script:IsUserMode) { return }
    if (-not (Test-SafeNpmCommand -Command $script:NpmExe)) {
        Write-Status 'warn' "npm 命令不安全，跳过 npm 前缀设置：$($script:NpmExe)"
        return
    }
    $npmPrefix = Join-Path $env:APPDATA 'npm'
    & $script:NpmExe config set prefix $npmPrefix 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Status 'warn' "npm 前缀设置失败（退出码 $LASTEXITCODE）"
        return
    }
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    if ($userPath -notlike "*$npmPrefix*") {
        [Environment]::SetEnvironmentVariable('PATH', "$userPath;$npmPrefix", 'User')
    }
    if ($env:Path -notlike "*$npmPrefix*") { $env:Path = "$env:Path;$npmPrefix" }
}

function Get-RunningOpenCodeProcesses {
    try {
        return @(Get-Process -Name opencode -ErrorAction SilentlyContinue)
    } catch {
        return @()
    }
}

function Resolve-OpenCodeExecutable {
    foreach ($name in @('opencode.cmd', 'opencode.exe')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.CommandType -eq 'Application') { return [string]$command.Source }
    }
    $command = Get-Command opencode -ErrorAction SilentlyContinue
    if ($command -and $command.CommandType -eq 'Application') { return [string]$command.Source }
    return $null
}

function Invoke-ToolStep {
    param(
        [string]$Label,
        [string]$FilePath,
        [string[]]$Arguments,
        [int]$TimeoutSec = 300,
        [switch]$Quiet
    )

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        Write-Status 'error' "$Label：未找到可执行命令"
        return $false
    }

    if (-not $Quiet) { Write-Status 'info' "$Label（最长等待 $([int]($TimeoutSec / 60)) 分钟）..." }
    try {
        $result = Invoke-CommandWithTimeout -FilePath $FilePath -Arguments $Arguments -TimeoutSec $TimeoutSec -HeartbeatSec 15 -OnHeartbeat {
            param($seconds)
            Write-Host ("      仍在执行... {0}s" -f $seconds) -ForegroundColor DarkGray
        }
    } catch {
        Write-Status 'error' "$Label 无法启动：$($_.Exception.Message)"
        return $false
    }

    if ($result.TimedOut) {
        Write-Status 'error' "$Label 超时（${TimeoutSec}s），已终止"
        return $false
    }
    if ($result.ExitCode -ne 0) {
        Write-Status 'error' "$Label 失败（退出码 $($result.ExitCode)）"
        $details = (@($result.Error, $result.Output) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n"
        if ($details) {
            foreach ($line in ($details -split "`r?`n" | Select-Object -Last 15)) {
                if (-not [string]::IsNullOrWhiteSpace($line)) { Write-Host "      $line" -ForegroundColor DarkGray }
            }
        }
        return $false
    }

    if (-not $Quiet) { Write-Status 'ok' "$Label 完成" }
    return $true
}

function Invoke-NpmStep {
    param(
        [string]$Label,
        [string[]]$Arguments,
        [int]$TimeoutSec = 300
    )

    if (-not (Test-SafeNpmCommand -Command $script:NpmExe)) {
        Write-Status 'error' "npm 命令不安全，无法执行：$Label"
        return $false
    }

    return Invoke-ToolStep -Label $Label -FilePath $script:NpmExe -Arguments $Arguments -TimeoutSec $TimeoutSec
}

function Invoke-PipStep {
    param(
        [string]$Label,
        [string]$PythonExe,
        [string[]]$Arguments,
        [int]$TimeoutSec = 600,
        [switch]$Quiet
    )

    return Invoke-ToolStep -Label $Label -FilePath $PythonExe -Arguments $Arguments -TimeoutSec $TimeoutSec -Quiet:$Quiet
}

function Invoke-OpenCodeSetup {
    param([hashtable]$EnvInfo)

    $plan = Get-OpenCodeInstallPlan -State @{
        Version       = $EnvInfo.OpenCodeVersion
        Path          = $EnvInfo.OpenCodePath
        NpmPrefix     = $EnvInfo.NpmPrefix
        HasV1         = $EnvInfo.HasV1
        HasV2         = $EnvInfo.HasV2
        LatestVersion = $EnvInfo.OpenCodeLatestVersion
    }
    if ($plan.Action -eq 'None') { Write-Status 'ok' $plan.Reason; return $true }
    Write-Status 'info' $plan.Reason

    if ($script:DryRun) {
        foreach ($command in $plan.Commands) { Write-Status 'plan' ('执行：' + ($command -join ' ')) }
        return $true
    }

    $running = @(Get-RunningOpenCodeProcesses)
    if ($running.Count -gt 0) {
        Write-Status 'warn' "检测到 OpenCode 正在运行（PID $($running.Id -join ', ')）"
        Write-Status 'info' 'Windows 无法替换正在运行的程序；请关闭所有 OpenCode 窗口后重新运行安装器。'
        Write-Status 'info' '本次跳过 OpenCode 更新，其他组件仍会继续安装。'
        return $true
    }

    Set-NpmMirror
    Set-NpmUserPrefix

    $npmInstallArgs = @('install', '-g', '@opencode/cli@latest', '--no-audit', '--no-fund', '--loglevel=error')

    if ($plan.Action -eq 'MigrateV1') {
        Write-Status 'info' '卸载 V1 包 opencode-ai...'
        if (-not (Invoke-NpmStep -Label '卸载 opencode-ai' -Arguments @('uninstall', '-g', 'opencode-ai') -TimeoutSec 120)) {
            Write-Status 'warn' 'V1 包卸载失败，继续安装 V2'
        }
        if (-not (Invoke-NpmStep -Label '安装 @opencode/cli' -Arguments $npmInstallArgs)) { return $false }
    } elseif ($plan.Action -eq 'UpdateNpm') {
        if (-not (Invoke-NpmStep -Label '更新 @opencode/cli' -Arguments $npmInstallArgs)) { return $false }
    } elseif ($plan.Action -eq 'UpdateCurl') {
        $opencodeExe = Resolve-OpenCodeExecutable
        if (-not $opencodeExe) { Write-Status 'error' '未找到 opencode 命令，无法执行升级'; return $false }
        Write-Status 'info' '执行 opencode upgrade（最长等待 5 分钟）...'
        $result = Invoke-CommandWithTimeout -FilePath $opencodeExe -Arguments @('upgrade') -TimeoutSec 300 -HeartbeatSec 15 -OnHeartbeat {
            param($seconds)
            Write-Host ("      仍在执行... {0}s" -f $seconds) -ForegroundColor DarkGray
        }
        if ($result.TimedOut) { Write-Status 'error' 'opencode upgrade 超时（300s），已终止'; return $false }
        if ($result.ExitCode -ne 0) {
            Write-Status 'error' "opencode upgrade 失败（退出码 $($result.ExitCode)）"
            $details = (@($result.Error, $result.Output) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "`n"
            if ($details) {
                foreach ($line in ($details -split "`r?`n" | Select-Object -Last 15)) {
                    if (-not [string]::IsNullOrWhiteSpace($line)) { Write-Host "      $line" -ForegroundColor DarkGray }
                }
            }
            return $false
        }
    } else {
        if (-not (Invoke-NpmStep -Label '安装 @opencode/cli' -Arguments $npmInstallArgs)) { return $false }
    }

    Update-ProcessPath
    $version = Get-OpenCodeVersion
    if ($version) {
        Write-Status 'ok' "OpenCode 已就绪（$version）"
        $opencodeExe = Resolve-OpenCodeExecutable
        if ($opencodeExe) { try { Invoke-CommandWithTimeout -FilePath $opencodeExe -Arguments @('service', 'restart') -TimeoutSec 30 -HeartbeatSec 0 | Out-Null } catch { $null = $_ } }
        return $true
    }
    Write-Status 'warn' 'OpenCode 命令暂未生效，请重开终端后验证'
    return $true
}

# ============================================================
# AI 配置
# ============================================================

function Save-ApiKey {
    param([string]$Path, [string]$ApiKey)

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($Path, $ApiKey, (New-Object System.Text.UTF8Encoding($false)))
    try { (Get-Item -LiteralPath $Path -Force).Attributes = 'Hidden' } catch { $null = $_ }
}

function Get-ProviderKeyReference {
    param([string]$ConfigJson, [string]$ProviderId)

    if ([string]::IsNullOrWhiteSpace($ConfigJson)) { return $null }
    try { $config = $ConfigJson | ConvertFrom-Json } catch { return $null }
    if ($config.providers) {
        $property = $config.providers.PSObject.Properties[$ProviderId]
        if ($property -and $property.Value.settings) { return [string]$property.Value.settings.apiKey }
    }
    if ($config.provider) {
        $property = $config.provider.PSObject.Properties[$ProviderId]
        if ($property -and $property.Value.options) { return [string]$property.Value.options.apiKey }
    }
    return $null
}

function Get-ProviderModelsFromApi {
    param([string]$BaseUrl, [string]$ApiKey, [string]$ProviderName)

    $url = Get-ApiModelsUrl -BaseUrl $BaseUrl
    $headers = @{}
    if (-not [string]::IsNullOrWhiteSpace($ApiKey)) { $headers['Authorization'] = "Bearer $ApiKey" }
    try {
        $response = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 30
        $json = $response | ConvertTo-Json -Depth 20 -Compress
        $models = @(ConvertFrom-ModelsApiResponse -Json $json)
        if ($models.Count -eq 0) { throw '接口未返回模型列表' }
        Write-Status 'ok' "$ProviderName：已获取 $($models.Count) 个模型"
        return $models
    } catch {
        $status = 0
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
            $status = [int]$_.Exception.Response.StatusCode
        }
        if ($status -eq 404) {
            Write-Status 'warn' '接口不支持 /models（404），请手动输入模型 ID'
        } else {
            Write-Status 'error' (Resolve-ApiFailureMessage -StatusCode $status -Endpoint $url)
        }
        return @()
    }
}

function Read-DeepSeekConfig {
    param([string]$ConfigJson)

    if ($script:DryRun) {
        Write-Status 'plan' "DryRun：调用 $((Get-ApiModelsUrl -BaseUrl $DEEPSEEK_BASE_URL)) 动态获取模型"
        return $null
    }

    $existingRef = Get-ProviderKeyReference -ConfigJson $ConfigJson -ProviderId 'deepseek'
    $existingKey = Resolve-ConfigValueReference -Value $existingRef -ConfigDir $script:ConfigRoot
    $apiKey = ''
    if ($existingKey) {
        Write-Status 'ok' '检测到已保存的 DeepSeek API Key'
        if (Confirm-YesNo '是否继续使用这个 Key？' $true) { $apiKey = $existingKey }
    }

    while ([string]::IsNullOrWhiteSpace($apiKey)) {
        $apiKey = Read-SecretText '请输入 DeepSeek API Key（sk- 开头，输入不显示）'
        if ([string]::IsNullOrWhiteSpace($apiKey)) { Write-Status 'error' 'API Key 不能为空' }
        else { $apiKey = $apiKey.Trim() }
    }

    $models = @(Get-ProviderModelsFromApi -BaseUrl $DEEPSEEK_BASE_URL -ApiKey $apiKey -ProviderName 'DeepSeek')
    if ($models.Count -eq 0) {
        Write-Status 'warn' '未能从 DeepSeek 获取模型列表'
        if (-not (Confirm-YesNo '仍要保存这个 Key 并手动指定模型吗？' $false)) { return $null }
        $manual = Read-Host '  手动输入模型 ID（逗号分隔，直接回车跳过 AI 配置）'
        if ([string]::IsNullOrWhiteSpace($manual)) { return $null }
        $models = @($manual -split '[,，\s]+' | Where-Object { $_ })
    }
    if ($models.Count -eq 0) { return $null }

    $defaultModel = Select-DefaultModel -Models $models
    $mainModel = $defaultModel
    if ($models.Count -gt 1) {
        Write-Host ''
        Write-Host '  可用模型：' -ForegroundColor Gray
        for ($i = 0; $i -lt $models.Count; $i++) {
            $mark = if ($models[$i] -eq $defaultModel) { '（推荐）' } else { '' }
            Write-Host ("      [{0}] {1} {2}" -f ($i + 1), $models[$i], $mark)
        }
        $choice = ConvertTo-MenuInput -Text (Read-Host "  选择默认模型 [1-$($models.Count)，回车使用推荐]")
        if (-not [string]::IsNullOrWhiteSpace($choice) -and $choice -match '^\d+$') {
            $index = [int]$choice
            if ($index -ge 1 -and $index -le $models.Count) { $mainModel = $models[$index - 1] }
        }
    }

    $titleModel = Select-DefaultModel -Models $models
    $keyPath = Join-Path $script:SecretsDir 'deepseek-api-key'
    Save-ApiKey -Path $keyPath -ApiKey $apiKey
    $apiKeyRef = Get-ApiKeyReference -Path $keyPath
    $style = Get-ConfigSectionStyle -ConfigJson $ConfigJson -Section 'provider'
    $patch = New-ProviderPatch -Style $style -ProviderId 'deepseek' -Name 'DeepSeek' -BaseUrl $DEEPSEEK_BASE_URL `
        -ApiKeyRef $apiKeyRef -Models $models -MainModel $mainModel -TitleModel $titleModel
    return @{
        Patch      = $patch
        Provider   = 'DeepSeek'
        MainModel  = $mainModel
        TitleModel = $titleModel
        Models     = $models
    }
}

function Read-CustomProviderConfig {
    param([string]$ConfigJson)

    if ($script:DryRun) {
        Write-Status 'plan' 'DryRun：跳过自定义 Provider 输入'
        return $null
    }

    Write-Host ''
    $providerId = ConvertTo-MenuInput -Text (Read-Host '  Provider ID（英文，如 siliconflow）')
    $providerId = $providerId.ToLowerInvariant()
    if ($providerId -notmatch '^[a-z0-9][a-z0-9-_]*$') {
        Write-Status 'error' 'Provider ID 格式不正确（只能包含小写字母、数字、- 和 _）'
        return $null
    }
    $providerName = ([string](Read-Host '  显示名称（中文，如 硅基流动）')).Trim()
    if ([string]::IsNullOrWhiteSpace($providerName)) { $providerName = $providerId }
    $baseUrl = ConvertTo-MenuInput -Text (Read-Host '  API 地址（如 https://api.siliconflow.cn/v1）')
    if ([string]::IsNullOrWhiteSpace($baseUrl)) {
        Write-Status 'error' 'API 地址不能为空'
        return $null
    }
    $baseUrl = $baseUrl.TrimEnd('/')
    $apiKey = Read-SecretText 'API Key（可留空，输入不显示）'
    if ($apiKey) { $apiKey = $apiKey.Trim() }

    $models = @(Get-ProviderModelsFromApi -BaseUrl $baseUrl -ApiKey $apiKey -ProviderName $providerName)
    if ($models.Count -eq 0) {
        $manual = Read-Host '  手动输入模型 ID（逗号分隔，直接回车取消）'
        if ([string]::IsNullOrWhiteSpace($manual)) { return $null }
        $models = @($manual -split '[,，\s]+' | Where-Object { $_ })
    }
    if ($models.Count -eq 0) { return $null }

    $defaultModel = Select-DefaultModel -Models $models
    $mainModel = $defaultModel
    if ($models.Count -gt 1) {
        Write-Host ''
        for ($i = 0; $i -lt $models.Count; $i++) {
            $mark = if ($models[$i] -eq $defaultModel) { '（推荐）' } else { '' }
            Write-Host ("      [{0}] {1} {2}" -f ($i + 1), $models[$i], $mark)
        }
        $choice = ConvertTo-MenuInput -Text (Read-Host "  选择默认模型 [1-$($models.Count)，回车使用推荐]")
        if (-not [string]::IsNullOrWhiteSpace($choice) -and $choice -match '^\d+$') {
            $index = [int]$choice
            if ($index -ge 1 -and $index -le $models.Count) { $mainModel = $models[$index - 1] }
        }
    }

    $titleModel = Select-DefaultModel -Models $models
    $apiKeyRef = ''
    if ($apiKey) {
        $keyPath = Join-Path $script:SecretsDir ("$providerId-api-key")
        Save-ApiKey -Path $keyPath -ApiKey $apiKey
        $apiKeyRef = Get-ApiKeyReference -Path $keyPath
    }
    $style = Get-ConfigSectionStyle -ConfigJson $ConfigJson -Section 'provider'
    $patch = New-ProviderPatch -Style $style -ProviderId $providerId -Name $providerName -BaseUrl $baseUrl `
        -ApiKeyRef $apiKeyRef -Models $models -MainModel $mainModel -TitleModel $titleModel
    return @{
        Patch      = $patch
        Provider   = $providerName
        MainModel  = $mainModel
        TitleModel = $titleModel
        Models     = $models
    }
}

# ============================================================
# 组件部署
# ============================================================

function Get-InstalledSkillVersion {
    param([string]$SkillName)

    $manifest = Read-InstallerManifest -Path $script:ManifestPath
    if ($manifest -and $manifest.skills.ContainsKey($SkillName)) { return [string]$manifest.skills[$SkillName] }
    return Get-SkillVersionFromSkillMd -Path (Join-Path (Join-Path $script:SkillsDir $SkillName) 'SKILL.md')
}

function Save-InstalledSkillVersion {
    param([string]$SkillName, [string]$Version)

    if ($script:DryRun) { return }
    $manifest = Read-InstallerManifest -Path $script:ManifestPath
    if (-not $manifest) { $manifest = @{ installerVersion = $script:InstallerVersion; skills = @{} } }
    $manifest.installerVersion = $script:InstallerVersion
    $manifest.skills[$SkillName] = $Version
    Save-InstallerManifest -Path $script:ManifestPath -Manifest $manifest
}

function Install-SkillPack {
    param([hashtable]$Skill)

    $installedVersion = Get-InstalledSkillVersion -SkillName $Skill.Name
    $action = Resolve-SkillAction -InstalledVersion $installedVersion -PayloadVersion $Skill.Version
    if ($action -eq 'Skip') {
        Write-Status 'ok' "$($Skill.Display)：已是最新版 v$($Skill.Version)"
        return $true
    }
    if ($script:DryRun) {
        $verb = if ($action -eq 'Install') { '安装' } else { "更新（$installedVersion → $($Skill.Version)）" }
        Write-Status 'plan' "${verb}技能包：$($Skill.Name)"
        return $true
    }

    Write-Status 'info' "$(if ($action -eq 'Install') { '正在安装' } else { '正在更新' }) $($Skill.Name)..."
    try {
        Expand-SkillPack -ZipPath (Join-Path $script:PayloadRoot $Skill.Zip) -InnerRoot $Skill.InnerRoot -Destination (Join-Path $script:SkillsDir $Skill.Name)
    } catch {
        Write-Status 'error' "$($Skill.Name) 部署失败：$($_.Exception.Message)"
        return $false
    }
    Write-Status 'ok' "$($Skill.Display) 已部署到 $($script:SkillsDir)\$($Skill.Name)"
    Save-InstalledSkillVersion -SkillName $Skill.Name -Version $Skill.Version
    return $true
}

function Install-SkillPythonDependencies {
    param([hashtable]$Skill)

    if (-not $Skill.NeedsPip) { return $true }
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        Write-Status 'warn' '未检测到 Python，跳过依赖安装（安装 Python 后重跑安装器即可补齐）'
        return $false
    }
    $requirements = Join-Path (Join-Path $script:SkillsDir $Skill.Name) 'requirements.txt'
    if (-not (Test-Path -LiteralPath $requirements)) { return $true }
    if ($script:DryRun) { Write-Status 'plan' "pip install -r $requirements（清华源）"; return $true }

    $pythonExe = if ($pythonCommand.Source) { [string]$pythonCommand.Source } else { 'python' }
    $pipArgs = Get-PipInstallArguments -RequirementsPath $requirements -IndexUrl $script:PipMirror
    if (-not (Invoke-PipStep -Label '安装 PPT Python 依赖' -PythonExe $pythonExe -Arguments $pipArgs)) {
        Write-Status 'warn' '镜像安装失败，尝试默认源...'
        $fallbackArgs = Get-PipInstallArguments -RequirementsPath $requirements -IndexUrl ''
        if (-not (Invoke-PipStep -Label '安装 PPT Python 依赖（默认源）' -PythonExe $pythonExe -Arguments $fallbackArgs)) {
            Write-Status 'warn' "依赖安装未完成，可稍后手动执行：python -m pip install -r `"$requirements`""
            return $false
        }
    }
    Write-Status 'ok' 'Python 依赖安装完成'
    return $true
}

function Get-McpPython {
    param([hashtable]$Mcp, [bool]$HasPython)

    if (-not $HasPython) { return $null }
    $venvDir = Get-McpVenvDir -McpRoot $script:McpDir -Key $Mcp.Key
    $venvPython = Join-Path $venvDir 'Scripts\python.exe'
    if (Test-Path -LiteralPath $venvPython) {
        try {
            $probe = Invoke-CommandWithTimeout -FilePath $venvPython -Arguments @('--version') -TimeoutSec 15 -HeartbeatSec 0
            if (-not $probe.TimedOut -and $probe.ExitCode -eq 0) { return $venvPython }
        } catch { $null = $_ }
        Remove-Item -LiteralPath $venvDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($script:DryRun) {
        Write-Status 'plan' "创建 $($Mcp.Name) 独立环境：$venvDir"
        return $venvPython
    }
    try {
        $pythonCommand = Get-Command python -ErrorAction Stop
        $pythonExe = if ($pythonCommand.Source) { [string]$pythonCommand.Source } else { 'python' }
        if (-not (Invoke-PipStep -Label "创建 $($Mcp.Name) 独立 Python 环境" -PythonExe $pythonExe -Arguments @('-m', 'venv', $venvDir) -TimeoutSec 180)) {
            throw 'python -m venv 失败'
        }
        return $venvPython
    } catch {
        Write-Status 'warn' "$($Mcp.Name) 独立环境创建失败，回退全局 Python：$($_.Exception.Message)"
        return 'python'
    }
}

function Install-McpPackage {
    param([hashtable]$Mcp, [string]$PythonExe)

    if (-not $PythonExe) { Write-Status 'warn' "未检测到 Python，跳过 $($Mcp.Display)"; return $false }
    if ($script:DryRun) {
        Write-Status 'plan' "pip install $($Mcp.PackageDir)（独立环境）"
        return $true
    }

    $packageRoot = Join-Path $script:PayloadRoot $Mcp.PackageDir
    try {
        # 有 zip 时必须重新解压，否则会复用 %TEMP% 里旧版本的源码
        if (-not (Expand-McpPackage -ZipPath (Join-Path $script:PayloadRoot $Mcp.Zip) -PackageRoot $packageRoot)) {
            Write-Status 'error' "缺少 $($Mcp.PackageDir) 源码包"
            return $false
        }
    } catch {
        Write-Status 'error' "解压 $($Mcp.Zip) 失败：$($_.Exception.Message)"
        return $false
    }

    $installArgs = @('-m', 'pip', 'install', '--upgrade', $packageRoot, '-i', $script:PipMirror)
    if (-not (Invoke-PipStep -Label "安装 $($Mcp.Display)" -PythonExe $PythonExe -Arguments $installArgs)) {
        Write-Status 'warn' '镜像安装失败，尝试默认源...'
        if (-not (Invoke-PipStep -Label "安装 $($Mcp.Display)（默认源）" -PythonExe $PythonExe -Arguments @('-m', 'pip', 'install', '--upgrade', $packageRoot))) {
            Write-Status 'error' "$($Mcp.Display) 安装失败"
            return $false
        }
    }

    # 同版本重跑时 pip 可能跳过安装；强制刷新本地包代码（不重装依赖）
    if (-not (Invoke-PipStep -Label "刷新 $($Mcp.Display) 代码" -PythonExe $PythonExe -Arguments @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', '--no-deps', $packageRoot) -Quiet)) {
        Write-Status 'warn' "$($Mcp.Display) 代码刷新失败，当前可能是旧版本"
    }

    # 只导入包本身会漏掉真正启动失败的 server 模块；这里导入实际入口
    $probeModule = [string]$Mcp.ProbeModule
    if ([string]::IsNullOrWhiteSpace($probeModule)) {
        Write-Status 'error' "$($Mcp.Display) 缺少启动探针模块"
        return $false
    }
    $probeArgs = @('-c', "import $probeModule")
    if (-not (Invoke-PipStep -Label "启动检查 $($Mcp.Display)" -PythonExe $PythonExe -Arguments $probeArgs -TimeoutSec 60 -Quiet)) {
        Write-Status 'warn' "$($Mcp.Display) 启动检查失败，尝试强制重装依赖..."
        $repairArgs = @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', $packageRoot, '-i', $script:PipMirror)
        if (-not (Invoke-PipStep -Label "修复 $($Mcp.Display) 依赖" -PythonExe $PythonExe -Arguments $repairArgs)) {
            Write-Status 'warn' '镜像修复失败，尝试默认源...'
            $repairArgs = @('-m', 'pip', 'install', '--upgrade', '--force-reinstall', $packageRoot)
            if (-not (Invoke-PipStep -Label "修复 $($Mcp.Display) 依赖（默认源）" -PythonExe $PythonExe -Arguments $repairArgs)) {
                Write-Status 'error' "$($Mcp.Display) 依赖修复失败"
                return $false
            }
        }
        if (-not (Invoke-PipStep -Label "复检 $($Mcp.Display)" -PythonExe $PythonExe -Arguments $probeArgs -TimeoutSec 60 -Quiet)) {
            Write-Status 'error' "$($Mcp.Display) 启动检查仍失败，已跳过配置（避免 OpenCode 报 MCP server failed）"
            return $false
        }
    }
    Write-Status 'ok' "$($Mcp.Display) 启动检查通过"
    return $true
}

function Install-OutlookMcp {
    param([string]$PythonExe)

    if (-not $PythonExe) {
        Write-Status 'warn' '未检测到 Python，跳过 Outlook 邮件助手'
        return $null
    }
    if (-not (Install-McpPackage -Mcp $script:OutlookMcp -PythonExe $PythonExe)) { return $null }
    return @{
        Command     = @($PythonExe, '-m', 'outlook_com_mcp')
        Environment = @{}
    }
}

function Install-ExcelMcp {
    param([string]$PythonExe)

    if (-not $PythonExe) {
        Write-Status 'warn' '未检测到 Python，跳过 Excel 表格助手'
        return $null
    }
    if (-not (Install-McpPackage -Mcp $script:ExcelMcp -PythonExe $PythonExe)) { return $null }
    return @{
        Command     = @($PythonExe, '-m', 'excel_mcp', 'stdio')
        Environment = @{}
    }
}

function Read-ComponentSelection {
    $items = @()
    foreach ($key in $script:EnabledComponents) {
        $items += @{
            Key      = $key
            Display  = $script:ComponentCatalog[$key]
            Selected = $true
        }
    }
    if ($script:DryRun) { return $items }

    while ($true) {
        Write-Host ''
        Write-Host '  请选择要安装的组件（输入序号切换勾选，a 全选，n 全不选，直接回车开始）:' -ForegroundColor Gray
        for ($i = 0; $i -lt $items.Count; $i++) {
            $mark = if ($items[$i].Selected) { '✓' } else { ' ' }
            Write-Host ("      [{0}] [{1}] {2}" -f ($i + 1), $mark, $items[$i].Display)
        }
        $answer = ConvertTo-MenuInput -Text (Read-Host '  你的选择')
        if ([string]::IsNullOrWhiteSpace($answer)) { return $items }
        if ($answer.ToLowerInvariant() -eq 'a') {
            $items | ForEach-Object { $_.Selected = $true }
            Write-Status 'ok' '已全选；再按一次回车开始安装'
            continue
        }
        if ($answer.ToLowerInvariant() -eq 'n') {
            $items | ForEach-Object { $_.Selected = $false }
            Write-Status 'info' '已全部取消；再按一次回车开始（将跳过组件）'
            continue
        }
        $changed = $false
        foreach ($token in ($answer -split '[^\d]+' | Where-Object { $_ })) {
            $index = [int]$token - 1
            if ($index -ge 0 -and $index -lt $items.Count) {
                $items[$index].Selected = -not $items[$index].Selected
                $changed = $true
            }
        }
        if (-not $changed) { Write-Status 'warn' '请输入序号、a、n，或直接回车开始' }
    }
}

function Install-Components {
    param([array]$Selection)

    $selected = @($Selection | Where-Object { $_.Selected } | ForEach-Object { $_.Key })
    $results = @()
    if ($selected.Count -eq 0) {
        Write-Status 'info' '未选择任何组件，跳过'
        return @{ Results = $results; ConfigJson = $script:ConfigJson }
    }

    $hasPython = $true
    if (-not $script:DryRun) { $hasPython = [bool](Get-PythonVersion) }
    if (-not $hasPython) { Write-Status 'warn' '未检测到 Python，PPT 依赖与 MCP 将被跳过' }

    $skillInstalled = $false
    if ($selected -contains 'ppt-master') {
        $ok = Install-SkillPack -Skill $script:PptMaster
        $depOk = $false
        if ($ok -and $hasPython) { $depOk = Install-SkillPythonDependencies -Skill $script:PptMaster }
        if ($ok) { $skillInstalled = $true }
        $detail = if ($depOk) { '技能 + Python 依赖' } elseif ($ok) { '技能已部署，Python 依赖未完成' } else { '部署失败' }
        $results += @{ Key = 'ppt-master'; Ok = $ok; Label = 'PPT-Master 6.7.3'; Detail = $detail }
    }
    if ($selected -contains 'guizang') {
        $ok = Install-SkillPack -Skill $script:GuiZang
        if ($ok) { $skillInstalled = $true }
        $results += @{ Key = 'guizang'; Ok = $ok; Label = '归藏网页 PPT'; Detail = if ($ok) { '已部署' } else { '部署失败' } }
    }

    if ($selected -contains 'outlook') {
        $outlookPython = Get-McpPython -Mcp $script:OutlookMcp -HasPython $hasPython
        $outlook = Install-OutlookMcp -PythonExe $outlookPython
        if ($outlook) {
            $mcpStyle = Get-ConfigSectionStyle -ConfigJson $script:ConfigJson -Section 'mcp'
            $patch = New-McpServerPatch -Style $mcpStyle -McpName 'outlook' -Command $outlook.Command -Environment $outlook.Environment
            $script:ConfigJson = Merge-OpenCodeConfigJson -ConfigJson $script:ConfigJson -PatchJson $patch
            $results += @{ Key = 'outlook'; Ok = $true; Label = 'Outlook 邮件助手'; Detail = '已配置' }
        } else {
            $results += @{ Key = 'outlook'; Ok = $false; Label = 'Outlook 邮件助手'; Detail = '安装失败，未更新配置' }
        }
    }
    if ($selected -contains 'excel') {
        $excelPython = Get-McpPython -Mcp $script:ExcelMcp -HasPython $hasPython
        $excel = Install-ExcelMcp -PythonExe $excelPython
        if ($excel) {
            $mcpStyle = Get-ConfigSectionStyle -ConfigJson $script:ConfigJson -Section 'mcp'
            $patch = New-McpServerPatch -Style $mcpStyle -McpName $script:ExcelMcp.Name -Command $excel.Command -Environment $excel.Environment
            $script:ConfigJson = Merge-OpenCodeConfigJson -ConfigJson $script:ConfigJson -PatchJson $patch
            $results += @{ Key = 'excel'; Ok = $true; Label = 'Excel 表格助手'; Detail = '已配置（my-excel）' }
        } else {
            $results += @{ Key = 'excel'; Ok = $false; Label = 'Excel 表格助手'; Detail = '安装失败，未更新配置' }
        }
    }

    if ($skillInstalled) {
        $permissionStyle = Get-EffectivePermissionStyle -ConfigJson $script:ConfigJson
        $script:ConfigJson = Add-SkillPermissionToJson -ConfigJson $script:ConfigJson -Style $permissionStyle
        Write-Status 'ok' '已允许 OpenCode 使用技能（permission）'
    }

    return @{ Results = $results; ConfigJson = $script:ConfigJson }
}

# ============================================================
# 配置写入与展示
# ============================================================

function Save-Config {
    param([string]$ConfigJson)

    if ([string]::IsNullOrWhiteSpace($ConfigJson)) { return }
    if ($ConfigJson -notmatch '"\$schema"') {
        $ConfigJson = Merge-OpenCodeConfigJson -ConfigJson $ConfigJson -PatchJson '{"$schema":"https://opencode.ai/config.json"}'
    }
    try { $null = $ConfigJson | ConvertFrom-Json } catch {
        Write-Status 'error' "配置 JSON 校验失败：$($_.Exception.Message)"
        throw
    }
    if ($script:DryRun) {
        Write-Status 'plan' "DryRun：写入配置 $($script:ConfigFile)"
        return
    }

    $dir = Split-Path -Parent $script:ConfigFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $script:ConfigFile) {
        $backup = "$($script:ConfigFile).backup-$(Get-Date -Format 'yyyyMMddHHmmss')"
        Copy-Item -LiteralPath $script:ConfigFile -Destination $backup -Force
        Write-Status 'info' "已备份原配置：$backup"
    }
    [System.IO.File]::WriteAllText($script:ConfigFile, $ConfigJson, (New-Object System.Text.UTF8Encoding($false)))
    Write-Status 'ok' "配置已保存：$($script:ConfigFile)"
}

function Get-EnvironmentInfo {
    $configExists = Test-Path -LiteralPath $script:ConfigFile
    $configValid = $true
    $configJson = ''
    if ($configExists) {
        try {
            $configJson = Get-Content -LiteralPath $script:ConfigFile -Raw -Encoding UTF8
            $null = $configJson | ConvertFrom-Json
        } catch {
            $configValid = $false
            $configJson = ''
        }
    }
    $script:ConfigJson = $configJson

    $providerStyle = 'V2'
    $mcpStyle = 'V2'
    $permissionStyle = 'V2'
    if ($configValid -and -not [string]::IsNullOrWhiteSpace($configJson)) {
        try {
            $providerStyle = Get-ConfigSectionStyle -ConfigJson $configJson -Section 'provider'
            $mcpStyle = Get-ConfigSectionStyle -ConfigJson $configJson -Section 'mcp'
            $permissionStyle = Get-ConfigSectionStyle -ConfigJson $configJson -Section 'permissions'
        } catch { $configValid = $false }
    }

    $packages = @(Get-NpmGlobalPackages)
    return @{
        OsCaption             = Get-OsCaption
        NodeVersion           = Get-NodeVersion
        PythonVersion         = Get-PythonVersion
        OpenCodeVersion       = Get-OpenCodeVersion
        OpenCodeLatestVersion = Get-OpenCodeLatestVersion -Registries @($script:NpmMirror, 'https://registry.npmjs.org')
        OpenCodePath          = Get-OpenCodePath
        NpmPrefix             = Get-NpmPrefix
        HasV1                 = ($packages -contains 'opencode-ai')
        HasV2                 = ($packages -contains '@opencode/cli')
        ConfigValid           = $configValid
        ProviderStyle         = $providerStyle
        McpStyle              = $mcpStyle
        PermissionStyle       = $permissionStyle
    }
}

function Show-NetworkAdvice {
    Write-Host ''
    Write-Host '  网络提示' -ForegroundColor Cyan
    Write-Host '  安装需要访问 npm / GitHub / 国内镜像；公司网络或代理可能拦截。' -ForegroundColor Yellow
    Write-Host '  建议先连接自己的手机热点再继续；如果公司网络/代理已确认可用，可忽略。' -ForegroundColor Yellow
}

function Show-Preflight {
    param([hashtable]$EnvInfo)

    Write-SectionTitle '环境预检'
    Write-KeyValue '系统' $EnvInfo.OsCaption 'ok'
    Write-KeyValue '运行模式' $(if ($script:IsAdmin) { '管理员' } else { '普通用户' }) 'ok'
    Write-KeyValue 'Node.js' $(if ($EnvInfo.NodeVersion) { $EnvInfo.NodeVersion } else { '未安装，将自动安装' }) $(if ($EnvInfo.NodeVersion) { 'ok' } else { 'warn' })
    Write-KeyValue 'Python' $(if ($EnvInfo.PythonVersion) { $EnvInfo.PythonVersion } else { '未安装，将自动安装' }) $(if ($EnvInfo.PythonVersion) { 'ok' } else { 'warn' })
    $pwshVersion = Get-PowerShell7Version
    if ($pwshVersion) {
        Write-KeyValue 'PowerShell 7' $pwshVersion 'ok'
    } elseif (Test-GitBashEnvironment) {
        Write-KeyValue 'PowerShell 7' '跳过（Git Bash）' 'ok'
    } else {
        Write-KeyValue 'PowerShell 7' '未安装，可选安装' 'warn'
    }
    if ($EnvInfo.OpenCodeVersion) {
        $method = if ($EnvInfo.HasV2) { 'npm' } elseif ($EnvInfo.OpenCodePath) { '系统安装' } else { '未知' }
        Write-KeyValue 'OpenCode' "v$($EnvInfo.OpenCodeVersion) · $method" 'ok'
    } else {
        Write-KeyValue 'OpenCode' '未安装，将安装 V2' 'warn'
    }
    $npmSafe = Test-SafeNpmCommand -Command $script:NpmExe
    Write-KeyValue 'npm 命令' $script:NpmExe $(if ($npmSafe) { 'ok' } else { 'warn' })
    if ($EnvInfo.ConfigValid) {
        Write-KeyValue '配置格式' "$($EnvInfo.ProviderStyle) provider / $($EnvInfo.McpStyle) mcp" 'ok'
    } else {
        Write-KeyValue '配置格式' '现有配置不是合法 JSON' 'error'
    }
    Write-KeyValue '安装目标' $script:ConfigRoot
}

function Show-PlanSummary {
    param([hashtable]$EnvInfo, [hashtable]$Plan)

    Write-SectionTitle '即将执行'
    if ($EnvInfo.NodeVersion) { Write-Status 'ok' "Node.js 已安装（$($EnvInfo.NodeVersion)）" }
    else { Write-Status 'plan' '安装 Node.js LTS' }
    if ($EnvInfo.PythonVersion) { Write-Status 'ok' "Python 已安装（$($EnvInfo.PythonVersion)）" }
    else { Write-Status 'plan' '安装 Python 3.12' }
    $pwshVersion = Get-PowerShell7Version
    if ($pwshVersion) { Write-Status 'ok' "PowerShell 7 已安装（$pwshVersion）" }
    elseif (Test-GitBashEnvironment) { Write-Status 'plan' 'PowerShell 7：跳过（检测到 Git Bash）' }
    else { Write-Status 'plan' "PowerShell 7：可选安装 $($script:PowerShell7Version)（用户目录，无需管理员）" }
    if ($Plan.Action -eq 'None') { Write-Status 'ok' "OpenCode：$($Plan.Reason)" }
    else { Write-Status 'plan' "OpenCode：$($Plan.Reason)" }
    Write-Status 'plan' 'DeepSeek：安装时调用 /models 动态获取模型'
    $names = @($script:EnabledComponents | ForEach-Object { $script:ComponentCatalog[$_] })
    Write-Status 'plan' ('组件：' + ($names -join '、'))
}

function Show-Completion {
    param([hashtable]$EnvInfo, $AiInfo)

    Write-CompletionBanner
    Write-SectionTitle '安装结果'
    if ($script:ComponentResults.Count -eq 0) {
        Write-Status 'info' '未安装组件'
    } else {
        foreach ($result in $script:ComponentResults) {
            if ($result.Ok) { Write-Status 'ok' "$($result.Label)：$($result.Detail)" }
            else { Write-Status 'warn' "$($result.Label)：$($result.Detail)" }
        }
    }
    if ($AiInfo) {
        Write-Status 'ok' "AI 模型：$($AiInfo.Provider) / $($AiInfo.MainModel)"
        Write-Status 'info' "标题模型：$($AiInfo.TitleModel)"
    } else {
        Write-Status 'info' 'AI 配置：保留现有配置或未配置'
    }
    $pwshVersion = Get-PowerShell7Version
    if ($pwshVersion) { Write-Status 'ok' "PowerShell 7：$pwshVersion" }
    elseif (Test-GitBashEnvironment) { Write-Status 'info' 'PowerShell 7：跳过（Git Bash）' }
    else { Write-Status 'info' 'PowerShell 7：未安装（可选）' }
    if ($script:DryRun) { Write-Status 'plan' '以上为 DryRun 计划，未实际修改系统' }

    Write-SectionTitle '路径'
    Write-KeyValue '配置文件' $script:ConfigFile
    Write-KeyValue '技能目录' $script:SkillsDir
    Write-KeyValue 'MCP 环境' $script:McpDir
    Write-KeyValue '日志文件' $script:LogFile

    if ($script:DryRun) {
        Write-Status 'plan' 'DryRun：将创建桌面快捷方式 OpenCode.lnk'
    } else {
        if (-not (Get-Command opencode -ErrorAction SilentlyContinue)) {
            Write-Status 'warn' '未找到 opencode，跳过桌面快捷方式'
        } else {
            $shortcutPath = New-OpenCodeShortcut
            if ($shortcutPath) {
                Write-Status 'ok' "桌面快捷方式：$shortcutPath"
            } else {
                Write-Status 'warn' '桌面快捷方式创建失败，可手动创建指向 opencode 的快捷方式'
            }
        }
    }

    Write-SectionTitle '下一步'
    Write-Host '    1. 打开终端，输入 opencode' -ForegroundColor White
    Write-Host '    2. 直接描述需求，例如“用这份 PDF 做一份 10 页 PPT”' -ForegroundColor Gray
    Write-Host '    3. 以后更新：重跑本安装器' -ForegroundColor Gray

    if ($script:DryRun) {
        Write-Status 'plan' 'DryRun：跳过启动 OpenCode'
    } elseif (Confirm-YesNo '是否立即启动 OpenCode？' $true) {
        if (-not (Get-Command opencode -ErrorAction SilentlyContinue)) {
            Write-Status 'warn' '未找到 opencode 命令，请手动在终端运行 opencode'
        } else {
            $pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
            if ($pwsh) {
                Start-Process $pwsh.Source -ArgumentList @('-NoExit', '-Command', 'opencode')
            } else {
                Start-Process cmd.exe -ArgumentList '/k opencode'
            }
        }
    }
}

# ============================================================
# 主流程
# ============================================================

function Main {
    try { Start-Transcript -Path $script:LogFile -Force | Out-Null } catch { $null = $_ }
    Write-AppBanner -Version $INSTALLER_VERSION -EditionInfo $editionInfo
    Write-Status 'info' "安装版本：$($editionInfo.Name) · 目标：$($script:ConfigRoot)"
    Write-Status 'info' "核心：$($script:InstallerCoreVersion) @ $($script:CorePath)"
    Write-Status 'info' "PowerShell：$($PSVersionTable.PSVersion) / npm：$($script:NpmExe)"
    if ($script:DryRun) { Write-Status 'info' 'DryRun 模式：只显示计划，不修改系统' }

    Write-Status 'info' '正在检查环境与 OpenCode 最新版...'
    $envInfo = Get-EnvironmentInfo
    Show-Preflight -EnvInfo $envInfo
    Show-NetworkAdvice

    $missingPayloads = @(Get-MissingPayloads -Root $script:PayloadRoot -Names $requiredPayloads)
    if ($missingPayloads.Count -gt 0) {
        if ($script:DryRun) {
            Write-Status 'warn' "DryRun：缺少内嵌资源 $($missingPayloads -join ', ')（开发模式属正常）"
        } else {
            Write-Status 'error' "安装器内嵌资源缺失：$($missingPayloads -join ', ')，请重新下载安装器"
            return 1
        }
    }

    if (-not $envInfo.ConfigValid) {
        Write-Status 'error' '现有 opencode.json 不是合法 JSON'
        if ($script:DryRun) {
            Write-Status 'plan' 'DryRun：跳过配置重建'
        } elseif (Confirm-YesNo '是否备份原文件并重建配置？' $false) {
            $backup = "$($script:ConfigFile).invalid-backup-$(Get-Date -Format 'yyyyMMddHHmmss')"
            Copy-Item -LiteralPath $script:ConfigFile -Destination $backup -Force
            Write-Status 'info' "已备份：$backup"
            $script:ConfigJson = ''
            $envInfo.ConfigValid = $true
        } else {
            Write-Status 'error' '已取消'
            return 1
        }
    }

    $plan = Get-OpenCodeInstallPlan -State @{
        Version       = $envInfo.OpenCodeVersion
        Path          = $envInfo.OpenCodePath
        NpmPrefix     = $envInfo.NpmPrefix
        HasV1         = $envInfo.HasV1
        HasV2         = $envInfo.HasV2
        LatestVersion = $envInfo.OpenCodeLatestVersion
    }
    Show-PlanSummary -EnvInfo $envInfo -Plan $plan
    if (-not $script:DryRun) {
        if (-not (Confirm-YesNo '按以上计划继续？' $true)) {
            Write-Status 'info' '已取消'
            return 0
        }
    }

    Write-StepHeader 1 8 'Node.js 环境'
    $nodeVersion = Get-NodeVersion
    if ($nodeVersion) {
        Write-Status 'ok' "已安装 $nodeVersion"
    } else {
        if (-not (Install-NodeJS)) {
            Write-Status 'error' 'Node.js 安装失败，无法继续'
            return 1
        }
    }

    Write-StepHeader 2 8 'Python 环境'
    $pythonVersion = Get-PythonVersion
    if ($pythonVersion) {
        Write-Status 'ok' "已安装 $pythonVersion"
    } elseif ($script:DryRun) {
        Write-Status 'plan' 'DryRun：将安装 Python 3.12'
    } elseif (Confirm-YesNo '是否安装 Python？' $true) {
        if (-not (Install-Python)) { Write-Status 'warn' 'Python 未安装，PPT 依赖与 MCP 将跳过' }
    } else {
        Write-Status 'skip' '已跳过 Python'
    }

    Write-StepHeader 3 8 'PowerShell 7'
    $powerShellOk = Install-PowerShell7
    if (-not $powerShellOk) { Write-Status 'warn' 'PowerShell 7 未安装，继续使用 Windows PowerShell 5.1' }

    Write-StepHeader 4 8 'OpenCode'
    $openCodeOk = Invoke-OpenCodeSetup -EnvInfo $envInfo
    if (-not $openCodeOk) { Write-Status 'warn' 'OpenCode 未就绪，后续配置仍会写入' }

    Write-StepHeader 5 8 'AI 模型'
    $aiInfo = $null
    $hasExistingProvider = $false
    if ($envInfo.ConfigValid -and -not [string]::IsNullOrWhiteSpace($script:ConfigJson)) {
        $hasExistingProvider = ($script:ConfigJson -match '"provider"\s*:') -or ($script:ConfigJson -match '"providers"\s*:')
    }
    $configureAi = $true
    if ($hasExistingProvider) {
        Write-Status 'info' '检测到已有 AI 配置'
        $configureAi = $script:DryRun -or (Confirm-YesNo '是否重新配置 AI 模型？' $false)
    }
    if ($configureAi) {
        if ($script:DryRun) {
            Write-Status 'plan' 'DryRun：配置 AI 模型（动态获取模型，交互跳过）'
        } else {
            Write-Host ''
            Write-Host '  请选择 AI 服务：' -ForegroundColor Gray
            Write-Host '      [1] DeepSeek（推荐 · 国内直连 · 动态获取模型）'
            Write-Host '      [2] 其他 OpenAI 兼容 API'
            $aiProvider = $null
            while (-not $aiProvider) {
                $choice = Read-Host '  请输入 1 或 2（直接回车 = 1）'
                $aiProvider = Resolve-AiProviderChoice -Choice $choice
                if (-not $aiProvider) { Write-Status 'warn' '请输入 1（DeepSeek）或 2（其他 OpenAI 兼容 API）' }
            }
            if ($aiProvider -eq 'deepseek') {
                Write-Status 'info' '已选择：DeepSeek'
                $aiInfo = Read-DeepSeekConfig -ConfigJson $script:ConfigJson
            } else {
                Write-Status 'info' '已选择：其他 OpenAI 兼容 API'
                $aiInfo = Read-CustomProviderConfig -ConfigJson $script:ConfigJson
            }
            if ($aiInfo -and $aiInfo.Patch) {
                $script:ConfigJson = Merge-OpenCodeConfigJson -ConfigJson $script:ConfigJson -PatchJson $aiInfo.Patch
                Write-Status 'ok' "AI 模型已配置：$($aiInfo.Provider) / $($aiInfo.MainModel)"
            }
        }
    } else {
        Write-Status 'ok' '保留现有 AI 配置'
    }

    Write-StepHeader 6 8 '组件'
    $selection = Read-ComponentSelection
    $componentResult = Install-Components -Selection $selection
    $script:ConfigJson = $componentResult.ConfigJson
    $script:ComponentResults = $componentResult.Results

    Write-StepHeader 7 8 '写入配置'
    if (-not $envInfo.ConfigValid) {
        Write-Status 'warn' '配置无效，未写入'
    } elseif (-not [string]::IsNullOrWhiteSpace($script:ConfigJson)) {
        Save-Config -ConfigJson $script:ConfigJson
    } else {
        Write-Status 'skip' '没有需要写入的配置'
    }

    Write-StepHeader 8 8 '完成'
    Show-Completion -EnvInfo $envInfo -AiInfo $aiInfo
    return 0
}

if ($SelfTest) {
    Write-Host ''
    Write-Host "  OpenCode 安装器自检 v$INSTALLER_VERSION" -ForegroundColor Cyan
    Write-Host '  ──────────────────────────────────────────────────────────' -ForegroundColor DarkGray
    $selfTestOk = Invoke-SelfTest
    exit $(if ($selfTestOk) { 0 } else { 1 })
}

$exitCode = 0
try {
    $exitCode = Main
} catch {
    $exitCode = 1
    Write-Host ''
    Write-Status 'error' "安装过程出现错误：$($_.Exception.Message)"
    Write-Status 'info' "详细日志：$($script:LogFile)"
} finally {
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    Wait-Key
}
exit $exitCode
