# ============================================================
# OpenCode 一键安装器 v3.0.0
# 外发版：内置 PPT-Master / 归藏网页 PPT 技能包 + 可选 Outlook MCP
# 适用于 Windows 10/11，简体中文，管理员/普通用户双模式
# ============================================================
[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$TestRoot = '',
    [string]$PayloadRoot = '',
    [switch]$NoPause
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { $null = $_ }
try { $OutputEncoding = [System.Text.Encoding]::UTF8 } catch { $null = $_ }
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch { $null = $_ }
$ErrorActionPreference = 'Continue'

$INSTALLER_VERSION = '3.0.0'
$PAYLOAD_DIR = Join-Path $env:TEMP 'opencode-installer'
$NPM_MIRROR = 'https://registry.npmmirror.com'
$PIP_MIRROR = 'https://pypi.tuna.tsinghua.edu.cn/simple'

$PPT_MASTER = @{
    Name = 'ppt-master'; Display = 'PPT-Master（原生 PPTX，PowerPoint 里可继续编辑）'
    Zip = 'ppt-master-skill-v6.6.0.zip'; InnerRoot = 'ppt-master/skills/ppt-master'
    Version = '6.6.0'; NeedsPip = $true
    Credit = 'PPT-Master by Hugo He · MIT · github.com/hugohe3/ppt-master'
}
$GUIZANG = @{
    Name = 'guizang-ppt-skill'; Display = '归藏网页 PPT（单文件 HTML 幻灯片，浏览器直接放映）'
    Zip = 'guizang-ppt-skill.zip'; InnerRoot = 'guizang-ppt-skill'
    Version = '2026.09.20'; NeedsPip = $false
    Credit = '归藏网页 PPT by 歸藏 · AGPL-3.0 · github.com/op7418/guizang-ppt-skill'
}
$OUTLOOK_MCP = @{
    Name = 'outlook'; Payload = 'outlook_mcp_server.py'
    Display = 'Outlook 邮件助手（MCP，需要桌面版 Outlook）'
    Packages = @('pywin32', 'fastmcp')
}
$REQUIRED_PAYLOADS = @($PPT_MASTER.Zip, $GUIZANG.Zip, $OUTLOOK_MCP.Payload)

$coreCandidates = @(Join-Path $PAYLOAD_DIR 'installer-core.ps1')
if ($PSScriptRoot) { $coreCandidates += (Join-Path $PSScriptRoot 'lib\installer-core.ps1') }
$corePath = ''
foreach ($candidate in $coreCandidates) {
    if (Test-Path -LiteralPath $candidate) { $corePath = $candidate; break }
}
if (-not $corePath) {
    Write-Host '  安装器组件缺失，请重新下载安装器' -ForegroundColor Red
    exit 1
}
. $corePath

if (-not $PayloadRoot) {
    $assetsDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'assets' } else { '' }
    $PayloadRoot = Get-PayloadRoot -TempRoot $PAYLOAD_DIR -AssetsDir $assetsDir -RequiredNames $REQUIRED_PAYLOADS
}

$script:IsAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$script:IsUserMode = -not $script:IsAdmin
$script:NoPause = $NoPause -or $DryRun

$configRoot = if ($TestRoot) { $TestRoot } else { Join-Path $env:USERPROFILE '.config\opencode' }
$skillsDir = Join-Path $configRoot 'skills'
$configFile = Join-Path $configRoot 'opencode.json'
$manifestPath = Join-Path $configRoot 'installer-manifest.json'
$outlookMcpDir = Join-Path $configRoot 'mcp-servers\outlook-mcp'

# ============================================================
# 界面输出
# ============================================================

function Clear-Screen {
    try { Clear-Host } catch { $null = $_ }
}

function Write-Banner {
    Clear-Screen
    $lines = @(
        '',
        '  ╔══════════════════════════════════════════════╗',
        '  ║        OpenCode 一键安装器  v3.0.0            ║',
        '  ║   内置 PPT 技能包 · Outlook 邮件助手          ║',
        '  ╚══════════════════════════════════════════════╝',
        '',
        '  OpenCode 是开源的 AI 编程助手，支持 DeepSeek 等模型。',
        '  本工具将帮助你完成安装、配置与技能部署。',
        ''
    )
    Write-Host ($lines -join "`n") -ForegroundColor Cyan
}
function Write-Step { param([string]$Message) Write-Host "  [√] $Message" -ForegroundColor Green }
function Write-Info { param([string]$Message) Write-Host "  [i] $Message" -ForegroundColor Yellow }
function Write-Warn { param([string]$Message) Write-Host "  [!] $Message" -ForegroundColor DarkYellow }
function Write-ErrorLine { param([string]$Message) Write-Host "  [×] $Message" -ForegroundColor Red }
function Write-Planned { param([string]$Message) Write-Host "  [计划] $Message" -ForegroundColor Magenta }
function Write-Section { param([string]$Title) Write-Host "  ──────── $Title ────────" -ForegroundColor Cyan }

function Wait-Key {
    if ($script:NoPause) { return }
    Write-Host ''
    Write-Host '  按任意键继续...' -ForegroundColor DarkGray
    try { $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { $null = Read-Host }
}

function Confirm-YesNo {
    param([string]$Question, [bool]$Default = $false)
    $suffix = if ($Default) { '(Y/n)' } else { '(y/N)' }
    while ($true) {
        $answer = Read-Host "  $Question $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        switch ($answer.Trim().ToLower()) {
            'y' { return $true }
            'n' { return $false }
            default { Write-Host '  请输入 y 或 n' -ForegroundColor Red }
        }
    }
}

# ============================================================
# 环境检测与安装
# ============================================================

function Test-WingetAvailable {
    try { $null = Get-Command winget -ErrorAction Stop; return $true } catch { return $false }
}

function Update-ProcessPath {
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    $machinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $env:Path = "$machinePath;$userPath;$env:Path"
}

function Test-NodeInstalled {
    try {
        $version = & node --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) { Write-Step "检测到 Node.js $version"; return $true }
    } catch { $null = $_ }
    return $false
}

function Install-NodeJS {
    Write-Info 'Node.js 未检测到，正在自动安装...'
    if ($DryRun) { Write-Planned '安装 Node.js LTS（winget 优先，失败则用国内镜像 MSI）'; return }

    if (Test-WingetAvailable) {
        Write-Info '通过 winget 安装 Node.js LTS...'
        $scope = if ($script:IsUserMode) { '--scope user ' } else { '' }
        $proc = Start-Process winget -ArgumentList "install OpenJS.NodeJS.LTS $scope--accept-source-agreements --accept-package-agreements --silent" -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            Update-ProcessPath
            if (Test-NodeInstalled) { Write-Step 'Node.js 安装完成（winget）'; return }
        }
        Write-Warn 'winget 方式未完成，改用镜像下载安装包...'
    }

    $indexJson = ''
    try { $indexJson = (Invoke-WebRequest -Uri 'https://npmmirror.com/mirrors/node/index.json' -UseBasicParsing -TimeoutSec 30).Content } catch { $null = $_ }
    $candidates = @(Get-NodeDownloadCandidates -IndexJson $indexJson)
    $msiPath = Join-Path $PAYLOAD_DIR 'nodejs-lts-x64.msi'
    $downloaded = $false
    foreach ($url in $candidates) {
        try {
            Write-Info "正在下载... $($url.Split('/')[-1])"
            Invoke-WebRequest -Uri $url -OutFile $msiPath -UseBasicParsing -TimeoutSec 600
            $downloaded = $true
            break
        } catch { Write-Info '此地址不可用，尝试备选...' }
    }
    if (-not $downloaded) {
        Write-ErrorLine 'Node.js 下载失败，请检查网络后重新运行'
        Write-Info '备用方案：手动安装 https://nodejs.org/zh-cn/ 后重新运行本程序'
        return
    }

    $msiArgs = if ($script:IsUserMode) { "/i `"$msiPath`" /quiet /norestart MSIINSTALLPERUSER=1" } else { "/i `"$msiPath`" /quiet /norestart" }
    Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -NoNewWindow
    Remove-Item -LiteralPath $msiPath -Force -ErrorAction SilentlyContinue
    Update-ProcessPath
    if (Test-NodeInstalled) { Write-Step 'Node.js 安装完成'; return }

    Write-Info 'Node.js 已安装，正在重新启动安装器以刷新环境...'
    Restart-Self
}

function Test-PythonInstalled {
    try {
        $version = & python --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) { Write-Step "检测到 $version"; return $true }
    } catch { $null = $_ }
    return $false
}

function Install-Python {
    Write-Info 'Python 未检测到，正在自动安装...'
    if ($DryRun) { Write-Planned '安装 Python 3.12（winget 优先，失败则用国内镜像安装包）'; return }

    if (Test-WingetAvailable) {
        Write-Info '通过 winget 安装 Python 3.12...'
        $proc = Start-Process winget -ArgumentList 'install Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent' -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            Update-ProcessPath
            if (Test-PythonInstalled) { Write-Step 'Python 安装完成（winget）'; return }
        }
        Write-Warn 'winget 方式未完成，改用镜像下载安装包...'
    }

    $urls = @(
        'https://npmmirror.com/mirrors/python/3.12.9/python-3.12.9-amd64.exe',
        'https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe'
    )
    $installerPath = Join-Path $PAYLOAD_DIR 'python-installer.exe'
    $downloaded = $false
    foreach ($url in $urls) {
        try {
            Write-Info '正在下载 Python...'
            Invoke-WebRequest -Uri $url -OutFile $installerPath -UseBasicParsing -TimeoutSec 600
            $downloaded = $true
            break
        } catch { Write-Info '此地址不可用，尝试备选...' }
    }
    if (-not $downloaded) {
        Write-ErrorLine 'Python 下载失败，请检查网络后重新运行'
        return
    }

    $scope = if ($script:IsUserMode) { 'InstallAllUsers=0' } else { 'InstallAllUsers=1' }
    Start-Process $installerPath -ArgumentList "/quiet $scope PrependPath=1 Include_pip=1" -Wait -NoNewWindow
    Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
    Update-ProcessPath
    if (Test-PythonInstalled) { Write-Step 'Python 安装完成'; return }

    Write-Info 'Python 已安装，正在重新启动安装器以刷新环境...'
    Restart-Self
}

function Restart-Self {
    if ($DryRun) { return }
    $exePath = ''
    try { $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { $null = $_ }
    if ($exePath -and $exePath.EndsWith('.exe')) {
        Write-Info '正在自动重启...'
        Start-Sleep 1
        Start-Process $exePath
    } else {
        Write-Info '请关闭本窗口后重新运行安装器'
    }
    exit 0
}

# ============================================================
# 镜像与 OpenCode
# ============================================================

function Set-NpmMirror {
    if (-not (Get-Command npm -ErrorAction SilentlyContinue)) { return }
    Write-Info '正在配置 npm 国内镜像...'
    if ($DryRun) { Write-Planned "npm config set registry $NPM_MIRROR"; return }
    & npm config set registry $NPM_MIRROR | Out-Null
    Write-Step 'npm 镜像已设置为 npmmirror.com'
}

function Set-PipMirror {
    if (-not (Get-Command pip -ErrorAction SilentlyContinue)) { return }
    Write-Info '正在配置 pip 国内镜像...'
    if ($DryRun) { Write-Planned "pip config set global.index-url $PIP_MIRROR"; return }
    & pip config set global.index-url $PIP_MIRROR | Out-Null
    Write-Step 'pip 镜像已设置为清华源'
}

function Set-NpmUserPrefix {
    if (-not $script:IsUserMode) { return }
    $npmPrefix = Join-Path $env:APPDATA 'npm'
    & npm config set prefix $npmPrefix 2>$null | Out-Null
    $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    if ($userPath -notlike "*$npmPrefix*") {
        [Environment]::SetEnvironmentVariable('PATH', "$userPath;$npmPrefix", 'User')
    }
    if ($env:Path -notlike "*$npmPrefix*") { $env:Path = "$env:Path;$npmPrefix" }
}

function Get-OpenCodeVersion {
    try {
        $version = & opencode --version 2>$null
        if ($LASTEXITCODE -eq 0 -and $version) { return [string]$version }
    } catch { $null = $_ }
    return $null
}

function Install-OpenCode {
    Set-NpmUserPrefix
    Write-Info '正在安装 OpenCode（npm install -g opencode-ai）...'
    if ($DryRun) { Write-Planned 'npm install -g opencode-ai'; return $true }
    & npm install -g opencode-ai 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Step 'OpenCode 安装完成'; return $true }
    Write-ErrorLine 'OpenCode 安装失败，请检查网络后重新运行'
    return $false
}

function Update-OpenCode {
    Write-Info '正在更新 OpenCode（opencode upgrade）...'
    if ($DryRun) { Write-Planned 'opencode upgrade（失败则回退 npm install -g opencode-ai@latest）'; return $true }
    & opencode upgrade 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Step "OpenCode 已更新到 $(Get-OpenCodeVersion)"; return $true }
    Write-Warn 'opencode upgrade 失败，改用 npm 更新...'
    & npm install -g opencode-ai@latest 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Step "OpenCode 已更新到 $(Get-OpenCodeVersion)"; return $true }
    Write-ErrorLine 'OpenCode 更新失败'
    return $false
}

# ============================================================
# API 配置
# ============================================================

function Test-DeepSeekKey {
    param([string]$ApiKey)
    Write-Info '正在检查 api.deepseek.com:443 连通性...'
    if (-not (Test-TcpEndpoint -HostName 'api.deepseek.com' -Port 443 -TimeoutMs 5000)) {
        Write-ErrorLine (Resolve-ApiFailureMessage -StatusCode 0 -Endpoint 'api.deepseek.com:443')
        return $false
    }

    Write-Info '正在验证 API Key...'
    try {
        $body = @{ model = 'deepseek-v4-pro'; messages = @(@{ role = 'user'; content = 'Hi' }); max_tokens = 10; stream = $false } | ConvertTo-Json -Depth 3 -Compress
        $response = Invoke-RestMethod -Uri 'https://api.deepseek.com/chat/completions' -Method Post -ContentType 'application/json' -Headers @{ Authorization = "Bearer $ApiKey" } -Body $body -TimeoutSec 15
        if ($response.choices) { Write-Step 'API Key 验证成功'; return $true }
    } catch {
        $statusCode = 0
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode) { $statusCode = [int]$_.Exception.Response.StatusCode }
        Write-ErrorLine (Resolve-ApiFailureMessage -StatusCode $statusCode -Endpoint 'api.deepseek.com:443')
    }
    return $false
}

function Save-OpenCodeConfigText {
    param([string]$Json)
    if (-not (Test-Path -LiteralPath $configRoot)) { New-Item -ItemType Directory -Force -Path $configRoot | Out-Null }
    [System.IO.File]::WriteAllText($configFile, $Json, (New-Object System.Text.UTF8Encoding($false)))
}

function Save-ConfigPatch {
    param([hashtable]$Patch)
    if ($DryRun) { Write-Planned "写入 opencode.json：$($Patch.Keys -join ', ')"; return }
    if (Test-Path -LiteralPath $configFile) {
        Copy-Item -LiteralPath $configFile -Destination "$configFile.backup-$(Get-Date -Format 'yyyyMMddHHmmss')" -Force
        $existing = Get-Content -LiteralPath $configFile -Raw -Encoding UTF8
    } else {
        $existing = ''
    }
    $merged = Merge-OpenCodeConfigJson -ConfigJson $existing -PatchJson ($Patch | ConvertTo-Json -Depth 100)
    Save-OpenCodeConfigText -Json $merged
    Write-Step "配置已保存：$configFile"
}

function Set-DeepSeekConfig {
    Write-Host ''
    Write-Host '  ┌──────────────────────────────────────────────┐' -ForegroundColor Yellow
    Write-Host '  │ 获取 DeepSeek API Key：                       │' -ForegroundColor Yellow
    Write-Host '  │   1. 打开 https://platform.deepseek.com       │' -ForegroundColor Yellow
    Write-Host '  │   2. 登录后点击左侧 API Keys                  │' -ForegroundColor Yellow
    Write-Host '  │   3. 创建新 Key 并复制                        │' -ForegroundColor Yellow
    Write-Host '  └──────────────────────────────────────────────┘' -ForegroundColor Yellow
    if (Confirm-YesNo '是否现在打开 DeepSeek 官网？' $true) { Start-Process 'https://platform.deepseek.com' }

    $apiKey = ''
    while ([string]::IsNullOrWhiteSpace($apiKey)) {
        $apiKey = Read-Host '  请输入 DeepSeek API Key（以 sk- 开头）'
        if ([string]::IsNullOrWhiteSpace($apiKey)) { Write-Host '  API Key 不能为空' -ForegroundColor Red }
    }
    if (-not (Test-DeepSeekKey -ApiKey $apiKey)) {
        if (-not (Confirm-YesNo 'API Key 验证未通过，仍要保存吗？' $false)) { return $false }
    }

    $patch = @{
        '$schema'     = 'https://opencode.ai/config.json'
        model         = 'deepseek/deepseek-v4-pro'
        small_model   = 'deepseek/deepseek-v4-flash'
        provider      = @{
            deepseek = @{
                npm     = '@ai-sdk/openai-compatible'
                name    = 'DeepSeek'
                options = @{ baseURL = 'https://api.deepseek.com/v1'; apiKey = $apiKey }
                models  = @{
                    'deepseek-v4-pro'   = @{ name = 'DeepSeek V4 Pro（推荐）' }
                    'deepseek-v4-flash' = @{ name = 'DeepSeek V4 Flash（快速）' }
                }
            }
        }
    }
    Save-ConfigPatch -Patch $patch
    return $true
}

function Set-CustomProviderConfig {
    Write-Host ''
    Write-Info '自定义 OpenAI 兼容 API（硅基流动 / 通义 / OpenAI / Ollama 均可）'
    $providerId = Read-Host '  Provider ID（英文，如 siliconflow）'
    $providerName = Read-Host '  显示名称（中文，如 硅基流动）'
    $baseUrl = Read-Host '  API 地址（如 https://api.siliconflow.cn/v1）'
    $bigModel = Read-Host '  大模型 ID'
    $smallModel = Read-Host '  小模型 ID（可留空复用大模型）'
    $apiKey = Read-Host '  API Key'
    if ([string]::IsNullOrWhiteSpace($providerId) -or [string]::IsNullOrWhiteSpace($baseUrl)) {
        Write-ErrorLine 'Provider ID 和 API 地址不能为空'
        return $false
    }

    $endpoint = Get-HostPortFromUrl -Url $baseUrl
    if (-not $endpoint) {
        Write-Warn "API 地址无法解析（仍会保存，请确认格式如 https://api.xxx.com/v1）"
    } else {
        Write-Info "正在检查 $($endpoint.Host):$($endpoint.Port) 连通性..."
        if (Test-TcpEndpoint -HostName $endpoint.Host -Port $endpoint.Port -TimeoutMs 3000) {
            Write-Step "API 地址可达（$($endpoint.Host):$($endpoint.Port)）"
        } else {
            Write-Warn (Resolve-ApiFailureMessage -StatusCode 0 -Endpoint "$($endpoint.Host):$($endpoint.Port)")
            if (-not (Confirm-YesNo '仍要保存此配置吗？' $false)) { return $false }
        }
    }

    $options = @{ baseURL = $baseUrl }
    if (-not [string]::IsNullOrWhiteSpace($apiKey)) { $options.apiKey = $apiKey }
    $big = if ($bigModel) { $bigModel } else { 'default' }
    $small = if ($smallModel) { $smallModel } else { $big }
    $models = @{ $big = @{ name = "$providerName（大模型）" } }
    if ($small -ne $big) { $models[$small] = @{ name = "$providerName（小模型）" } }

    $patch = @{
        '$schema'   = 'https://opencode.ai/config.json'
        model       = "$providerId/$big"
        small_model = "$providerId/$small"
        provider    = @{ $providerId = @{ npm = '@ai-sdk/openai-compatible'; name = $providerName; options = $options; models = $models } }
    }
    Save-ConfigPatch -Patch $patch
    return $true
}

# ============================================================
# 组件部署（技能包 / MCP）
# ============================================================

function Get-InstalledSkillVersion {
    param([string]$SkillName)
    $manifest = Read-InstallerManifest -Path $manifestPath
    if ($manifest -and $manifest.skills.ContainsKey($SkillName)) { return [string]$manifest.skills[$SkillName] }
    return Get-SkillVersionFromSkillMd -Path (Join-Path (Join-Path $skillsDir $SkillName) 'SKILL.md')
}

function Save-InstalledSkillVersion {
    param([string]$SkillName, [string]$Version)
    if ($DryRun) { return }
    $manifest = Read-InstallerManifest -Path $manifestPath
    if (-not $manifest) { $manifest = @{ installerVersion = $INSTALLER_VERSION; skills = @{} } }
    $manifest.installerVersion = $INSTALLER_VERSION
    $manifest.skills[$SkillName] = $Version
    Save-InstallerManifest -Path $manifestPath -Manifest $manifest
}

function Install-SkillPack {
    param([hashtable]$Skill)

    $installedVersion = Get-InstalledSkillVersion -SkillName $Skill.Name
    $action = Resolve-SkillAction -InstalledVersion $installedVersion -PayloadVersion $Skill.Version
    if ($action -eq 'Skip') {
        Write-Step "$($Skill.Display)：已是最新版 v$($Skill.Version)，跳过"
        return $true
    }
    if ($DryRun) {
        Write-Planned "$(if ($action -eq 'Install') { '安装' } else { '更新（' + $installedVersion + ' → ' + $Skill.Version + '）' })技能包：$($Skill.Name)"
        return $true
    }

    Write-Info "$(if ($action -eq 'Install') { '正在安装' } else { '正在更新' }) $($Skill.Name) ..."
    try {
        Expand-SkillPack -ZipPath (Join-Path $PayloadRoot $Skill.Zip) -InnerRoot $Skill.InnerRoot -Destination (Join-Path $skillsDir $Skill.Name)
    } catch {
        Write-ErrorLine "$($Skill.Name) 部署失败：$($_.Exception.Message)"
        return $false
    }
    Write-Step "$($Skill.Display) 已部署到 $skillsDir\$($Skill.Name)"
    Save-InstalledSkillVersion -SkillName $Skill.Name -Version $Skill.Version
    return $true
}

function Install-SkillPythonDependencies {
    param([hashtable]$Skill)
    if (-not $Skill.NeedsPip) { return $true }
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        Write-Warn '未检测到 Python，跳过依赖安装（技能已部署，安装 Python 后重跑本安装器即可补齐）'
        return $false
    }
    $requirements = Join-Path (Join-Path $skillsDir $Skill.Name) 'requirements.txt'
    if (-not (Test-Path -LiteralPath $requirements)) { return $true }
    if ($DryRun) { Write-Planned "pip install -r $requirements（清华源）"; return $true }

    Write-Info '正在安装 Python 依赖（约几分钟，请耐心等待）...'
    $pipArgs = Get-PipInstallArguments -RequirementsPath $requirements -IndexUrl $PIP_MIRROR
    & python @pipArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Warn '镜像安装失败，尝试默认源...'
        $fallbackArgs = Get-PipInstallArguments -RequirementsPath $requirements -IndexUrl ''
        & python @fallbackArgs
    }
    if ($LASTEXITCODE -eq 0) { Write-Step 'Python 依赖安装完成'; return $true }
    Write-Warn "依赖安装未完成，可稍后手动执行：python -m pip install -r `"$requirements`""
    return $false
}

function Install-OutlookMcp {
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        Write-Warn '未检测到 Python，无法部署 Outlook 邮件助手（可安装 Python 后重跑本安装器）'
        return $false
    }
    if ($DryRun) {
        Write-Planned "部署 Outlook MCP 到 $outlookMcpDir 并安装 pywin32、fastmcp"
        Write-Planned '合并写入 opencode.json 的 mcp.outlook 配置'
        return $true
    }

    Write-Info '正在部署 Outlook 邮件助手...'
    New-Item -ItemType Directory -Force -Path $outlookMcpDir | Out-Null
    Copy-Item -LiteralPath (Join-Path $PayloadRoot $OUTLOOK_MCP.Payload) -Destination (Join-Path $outlookMcpDir $OUTLOOK_MCP.Payload) -Force

    $pipArgs = @('-m', 'pip', 'install') + $OUTLOOK_MCP.Packages + @('-i', $PIP_MIRROR)
    & python @pipArgs
    if ($LASTEXITCODE -ne 0) {
        $fallbackArgs = @('-m', 'pip', 'install') + $OUTLOOK_MCP.Packages
        & python @fallbackArgs
    }
    if ($LASTEXITCODE -eq 0) { Write-Step 'pywin32、fastmcp 安装完成' } else { Write-Warn '依赖安装未完成，MCP 可能无法启动' }

    $mcpConfig = Get-Content -LiteralPath $configFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    $merged = Update-OpenCodeConfigJson -ConfigJson $mcpConfig -McpName $OUTLOOK_MCP.Name -McpCommand @('python', (Join-Path $outlookMcpDir $OUTLOOK_MCP.Payload)) -AllowSkillPermission $false
    Save-OpenCodeConfigText -Json $merged
    Write-Step "Outlook 邮件助手已配置（$configFile）"
    return $true
}

function Read-ComponentSelection {
    $items = @(
        @{ Key = 'ppt-master'; Display = $PPT_MASTER.Display; Selected = $true },
        @{ Key = 'guizang'; Display = $GUIZANG.Display; Selected = $true },
        @{ Key = 'outlook'; Display = "$($OUTLOOK_MCP.Display)"; Selected = $true }
    )
    if ($DryRun) { return $items }

    while ($true) {
        Write-Host ''
        Write-Host '  请选择要安装的组件（输入序号切换勾选，a 全选，n 全不选，直接回车开始）:' -ForegroundColor Gray
        for ($i = 0; $i -lt $items.Count; $i++) {
            $mark = if ($items[$i].Selected) { '√' } else { ' ' }
            Write-Host ("      [$($i + 1)] [$mark] $($items[$i].Display)")
        }
        $answer = Read-Host '  你的选择'
        if ([string]::IsNullOrWhiteSpace($answer)) { return $items }
        if ($answer.Trim().ToLower() -eq 'a') { $items | ForEach-Object { $_.Selected = $true }; continue }
        if ($answer.Trim().ToLower() -eq 'n') { $items | ForEach-Object { $_.Selected = $false }; continue }
        foreach ($token in ($answer -split '[^\d]+' | Where-Object { $_ })) {
            $index = [int]$token - 1
            if ($index -ge 0 -and $index -lt $items.Count) { $items[$index].Selected = -not $items[$index].Selected }
        }
    }
}

function Install-Components {
    $selection = Read-ComponentSelection
    $selectedKeys = @($selection | Where-Object { $_.Selected } | ForEach-Object { $_.Key })
    if ($selectedKeys.Count -eq 0) {
        Write-Info '未选择任何组件，跳过'
        return
    }

    $anySkill = $selectedKeys -contains 'ppt-master' -or $selectedKeys -contains 'guizang'
    if ($selectedKeys -contains 'ppt-master') {
        Install-SkillPack -Skill $PPT_MASTER | Out-Null
        Install-SkillPythonDependencies -Skill $PPT_MASTER | Out-Null
    }
    if ($selectedKeys -contains 'guizang') { Install-SkillPack -Skill $GUIZANG | Out-Null }
    if ($selectedKeys -contains 'outlook') { Install-OutlookMcp | Out-Null }

    if ($anySkill -and -not $DryRun) {
        $existing = Get-Content -LiteralPath $configFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
        Save-OpenCodeConfigText -Json (Add-SkillPermissionJson -ConfigJson $existing)
        Write-Step '已允许 OpenCode 使用技能（permission.skill）'
    }
}

# ============================================================
# 完成
# ============================================================

function Show-Completion {
    Clear-Screen
    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════╗' -ForegroundColor Green
    Write-Host '  ║          OpenCode 安装完成！                  ║' -ForegroundColor Green
    Write-Host '  ╚══════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host ''
    Write-Host '  [使用方式]' -ForegroundColor Cyan
    Write-Host '    1. 打开终端（PowerShell 或 CMD）' -ForegroundColor White
    Write-Host '    2. 输入 opencode 并回车' -ForegroundColor White
    Write-Host '    3. 直接说需求，例如：' -ForegroundColor White
    Write-Host '       “用这份 PDF 做一份 10 页 PPT”        （走 PPT-Master）' -ForegroundColor Gray
    Write-Host '       “做一份分享风格的网页 PPT”           （走归藏网页 PPT）' -ForegroundColor Gray
    Write-Host '       “帮我找一下上周的邮件”               （走 Outlook 邮件助手）' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  [以后更新]' -ForegroundColor Cyan
    Write-Host '    OpenCode：重跑本安装器，或命令行执行 opencode upgrade' -ForegroundColor White
    Write-Host '    技能包：重跑新版安装器（按版本自动覆盖更新，.env 会保留）' -ForegroundColor White
    Write-Host ''
    Write-Host '  [说明]' -ForegroundColor Cyan
    Write-Host "    配置文件：$configFile" -ForegroundColor Gray
    Write-Host "    技能目录：$skillsDir" -ForegroundColor Gray
    Write-Host '    Outlook 邮件助手需要桌面版 Outlook（网页版邮箱不可用）' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  [技能来源与许可]' -ForegroundColor Cyan
    Write-Host "    $($PPT_MASTER.Credit)" -ForegroundColor DarkGray
    Write-Host "    $($GUIZANG.Credit)" -ForegroundColor Gray
    Write-Host ''
    if ($DryRun) {
        Write-Planned '询问是否立即启动 OpenCode（DryRun 跳过）'
    } elseif (Confirm-YesNo '是否立即启动 OpenCode？' $true) {
        if (Get-Command opencode -ErrorAction SilentlyContinue) { Start-Process cmd.exe -ArgumentList '/k opencode' }
        else { Write-Warn '未找到 opencode 命令，请手动在终端输入 opencode' }
    }
}

# ============================================================
# 主流程
# ============================================================

function Main {
    Write-Banner
    if ($DryRun) { Write-Info 'DryRun 模式：只显示计划，不修改系统' }
    if ($script:IsUserMode) {
        Write-Info '以普通用户模式运行（Node.js / Python 将安装到用户目录）'
    } else {
        Write-Info '以管理员模式运行'
    }
    Write-Info "安装目标：$configRoot"

    $missingPayloads = @(Get-MissingPayloads -Root $PayloadRoot -Names $REQUIRED_PAYLOADS)
    if ($missingPayloads.Count -gt 0) {
        if ($DryRun) { Write-Warn "DryRun：未找到内嵌资源 $($missingPayloads -join ', ')（开发模式属正常）" }
        else {
            Write-ErrorLine "安装器内嵌资源缺失：$($missingPayloads -join ', ')，请重新下载安装器"
            return
        }
    }

    Write-Section '第 1 步：Node.js 环境'
    if (-not (Test-NodeInstalled)) { Install-NodeJS }

    Write-Section '第 2 步：Python 环境（PPT 技能包需要）'
    if (-not (Test-PythonInstalled)) {
        if ($DryRun -or (Confirm-YesNo '是否安装 Python？' $true)) { Install-Python }
    }

    Write-Section '第 3 步：镜像与基础工具'
    Set-NpmMirror
    Set-PipMirror

    Write-Section '第 4 步：安装 / 更新 OpenCode'
    $currentVersion = Get-OpenCodeVersion
    if ($currentVersion) {
        Write-Step "OpenCode 已安装（$currentVersion）"
        if ($DryRun -or (Confirm-YesNo '是否更新到最新版？' $true)) { Update-OpenCode | Out-Null }
    } else {
        Install-OpenCode | Out-Null
    }

    Write-Section '第 5 步：配置 AI 模型'
    $hasConfig = (Test-Path -LiteralPath $configFile) -and ((Get-Content -LiteralPath $configFile -Raw -Encoding UTF8) -match '"provider"')
    $configureApi = $true
    if ($hasConfig) {
        Write-Info '检测到已有 AI 配置'
        $configureApi = $DryRun -or (Confirm-YesNo '是否重新配置 AI 模型？' $false)
    }
    if ($configureApi) {
        if ($DryRun) { Write-Planned '配置 AI 模型（交互式，DryRun 跳过）' }
        else {
            Write-Host ''
            Write-Host '  请选择你使用的 AI 服务：' -ForegroundColor Gray
            Write-Host '      [1] DeepSeek（推荐 · 国内直连）'
            Write-Host '      [2] 其他兼容 OpenAI 的 API'
            $choice = Read-Host '  请输入 1 或 2'
            if ($choice -eq '1') { Set-DeepSeekConfig | Out-Null } else { Set-CustomProviderConfig | Out-Null }
        }
    } else {
        Write-Step '保留现有 AI 配置'
    }

    Write-Section '第 6 步：部署技能与组件'
    Install-Components

    Write-Section '完成'
    Show-Completion
}

$exitCode = 0
try {
    Main
} catch {
    $exitCode = 1
    Write-Host ''
    Write-ErrorLine "安装过程出现错误：$_"
    Write-Info '请截图反馈，或手动执行：npm install -g opencode-ai'
} finally {
    if (-not $script:NoPause) {
        Write-Host ''
        Write-Host '  按任意键退出...' -ForegroundColor DarkGray
        try { $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { $null = Read-Host '  按 Enter 退出' }
    }
}
exit $exitCode
