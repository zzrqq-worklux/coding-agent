# ============================================================
# build.ps1 - 构建 OpenCode 安装器 exe
# 用法：
#   powershell -ExecutionPolicy Bypass -File build\build.ps1                # 构建（不签名）
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign         # 构建并用自签名证书签名
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -SkipTests     # 跳过单测（仅调试用）
# ============================================================
[CmdletBinding()]
param(
    [switch]$SkipTests,
    [switch]$NoSmokeTest,
    [switch]$Sign,
    [string]$PfxPath = '',
    [SecureString]$PfxPassword = $null,
    [string]$TimestampUrl = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root 'src'
$assetsDir = Join-Path $srcDir 'assets'
$distDir = Join-Path $PSScriptRoot 'dist'
$vendorPs2exe = Join-Path $PSScriptRoot 'vendor\ps2exe.ps1'
$installScript = Join-Path $srcDir 'install.ps1'
$coreScript = Join-Path $srcDir 'lib\installer-core.ps1'
$exePath = Join-Path $distDir 'OpenCode安装器.exe'
$guizangZip = Join-Path $distDir 'guizang-ppt-skill.zip'
$version = '3.0.0.0'
$requiredResources = @('install.ps1', 'installer-core.ps1', 'ppt-master-skill-v6.6.0.zip', 'guizang-ppt-skill.zip', 'outlook_mcp_server.py')

function Write-BuildStep { param([string]$Message) Write-Host "  [build] $Message" -ForegroundColor Cyan }

function Set-Utf8Bom {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($true)))
    Write-BuildStep "已补 UTF-8 BOM：$(Split-Path -Leaf $Path)"
}

function New-GuiZhangSkillZip {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $source = Join-Path $assetsDir 'guizang-ppt-skill'
    if (-not (Test-Path -LiteralPath $source)) { throw "缺少技能源目录：$source" }
    if (Test-Path -LiteralPath $guizangZip) { Remove-Item -LiteralPath $guizangZip -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($source, $guizangZip, [System.IO.Compression.CompressionLevel]::Optimal, $true)
    Write-BuildStep ("打包归藏技能：{0:N1} MB" -f ((Get-Item $guizangZip).Length / 1MB))
}

function Invoke-BuildTests {
    Write-BuildStep '运行单元测试...'
    $result = Invoke-Pester -Script (Join-Path $root 'tests\installer-core.Tests.ps1') -PassThru
    if (-not $result -or $result.FailedCount -gt 0) { throw "单元测试未通过（失败 $($result.FailedCount) 个），已中止构建" }
    Write-BuildStep "测试通过：$($result.PassedCount) 个"
}

function Test-EmbeddedResources {
    param([string]$ExePath, [string[]]$Names)
    # 在子进程中加载校验，避免当前进程锁定 exe 导致后续签名失败
    $probe = "([System.Reflection.Assembly]::LoadFile('$ExePath')).GetManifestResourceNames()"
    $resources = @(& powershell -NoProfile -Command $probe)
    if ($resources.Count -eq 0) { throw '内嵌资源校验失败：无法读取 exe 资源列表' }
    $missing = @($Names | Where-Object { $resources -notcontains $_ })
    if ($missing.Count -gt 0) { throw "exe 内嵌资源缺失：$($missing -join ', ')" }
    Write-BuildStep "内嵌资源校验通过（$($resources.Count) 项）"
}

function Get-BuildCertificate {
    param([string]$PfxPath, [SecureString]$PfxPassword)

    if ($PfxPath) {
        if (-not $PfxPassword) { $PfxPassword = Read-Host "  请输入 PFX 密码（$(Split-Path -Leaf $PfxPath)）" -AsSecureString }
        $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
        return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $PfxPassword, $flags)
    }

    $subject = 'CN=OpenCode Installer (Self-Signed)'
    $existing = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue | Where-Object { $_.Subject -eq $subject } | Select-Object -First 1
    if ($existing) { return $existing }

    Write-BuildStep '生成自签名代码签名证书（3 年有效期，存放 CurrentUser\My）'
    return New-SelfSignedCertificate -Type CodeSigningCert -Subject $subject -CertStoreLocation Cert:\CurrentUser\My -KeyUsage DigitalSignature -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(3)
}

function Invoke-SmokeTest {
    param([string]$ExePath)
    $smokeRoot = Join-Path $env:TEMP ('opencode-installer-smoke-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $smokeRoot | Out-Null
    try {
        $proc = Start-Process -FilePath $ExePath -ArgumentList @('-DryRun', '-TestRoot', $smokeRoot, '-NoPause') -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -ne 0) { throw "冒烟测试失败（退出码 $($proc.ExitCode)）" }
        Write-BuildStep '冒烟测试通过（-DryRun 全流程，退出码 0）'
    } finally {
        Remove-Item -LiteralPath $smokeRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
# 构建流程
# ============================================================

Write-Host ''
Write-Host '  ==== OpenCode 安装器 exe 构建 ====' -ForegroundColor Green
New-Item -ItemType Directory -Force -Path $distDir | Out-Null
foreach ($file in @($installScript, $coreScript)) { Set-Utf8Bom -Path $file }
if (-not $SkipTests) { Invoke-BuildTests } else { Write-BuildStep '跳过单元测试（-SkipTests）' }

New-GuiZhangSkillZip

$embedFiles = @{
    '%TEMP%\opencode-installer\ppt-master-skill-v6.6.0.zip' = (Join-Path $assetsDir 'ppt-master-skill-v6.6.0.zip')
    '%TEMP%\opencode-installer\guizang-ppt-skill.zip'       = $guizangZip
    '%TEMP%\opencode-installer\outlook_mcp_server.py'       = (Join-Path $assetsDir 'outlook_mcp_server.py')
    '%TEMP%\opencode-installer\installer-core.ps1'          = $coreScript
}
foreach ($source in $embedFiles.Values) {
    if (-not (Test-Path -LiteralPath $source)) { throw "缺少嵌入资源：$source" }
}

# 加载 ps2exe：必须在脚本顶层 dot-source（函数内 dot-source 在函数返回后即失效）
$ps2exeModule = Get-Module -ListAvailable ps2exe | Select-Object -First 1
if ($ps2exeModule) {
    Import-Module ps2exe -Force
    Write-BuildStep "使用 ps2exe 模块 v$($ps2exeModule.Version)"
} else {
    if (-not (Test-Path -LiteralPath $vendorPs2exe)) { throw '未找到 ps2exe（模块未安装且缺少 vendor\ps2exe.ps1）' }
    . $vendorPs2exe
    Write-BuildStep '使用 vendor\ps2exe.ps1（离线）'
}
if (-not (Get-Command Invoke-PS2EXE -ErrorAction SilentlyContinue)) { throw 'ps2exe 加载失败：未找到 Invoke-PS2EXE' }

Write-BuildStep '编译中（59MB 资源，可能需要几分钟）...'
if (Test-Path -LiteralPath $exePath) { Remove-Item -LiteralPath $exePath -Force }
Invoke-PS2EXE -inputFile $installScript -outputFile $exePath `
    -title 'OpenCode 安装器' `
    -description 'OpenCode 一键安装器：内置 PPT-Master、归藏网页 PPT 技能包与 Outlook 邮件助手' `
    -product 'OpenCode 安装器' `
    -version $version `
    -x64 `
    -embedFiles $embedFiles

if (-not (Test-Path -LiteralPath $exePath)) { throw '编译失败：未生成 exe' }
Write-BuildStep ("exe 已生成：{0:N1} MB" -f ((Get-Item $exePath).Length / 1MB))
Test-EmbeddedResources -ExePath $exePath -Names $requiredResources

if (-not $NoSmokeTest) { Invoke-SmokeTest -ExePath $exePath }

if ($Sign) {
    $cert = Get-BuildCertificate -PfxPath $PfxPath -PfxPassword $PfxPassword
    $signArgs = @{ FilePath = $exePath; Certificate = $cert; HashAlgorithm = 'SHA256' }
    if ($TimestampUrl) { $signArgs.TimestampServer = $TimestampUrl }
    $signature = Set-AuthenticodeSignature @signArgs
    if ($signature.SignerCertificate -and $signature.Status -eq 'UnknownError') {
        Write-BuildStep "签名已写入（发布者：$($signature.SignerCertificate.Subject)）；自签名根不受本机信任，状态 UnknownError 属预期"
    } else {
        Write-BuildStep "签名状态：$($signature.Status)"
    }
    Export-Certificate -Cert $cert -FilePath (Join-Path $distDir 'OpenCode安装器-签名证书.cer') | Out-Null
    Write-BuildStep '已导出公钥证书（.cer，供企业内信任导入）'
}

Write-Host ''
Write-Host '  构建完成' -ForegroundColor Green
Write-Host "  输出：$exePath"
if ($Sign) {
    Write-Host '  说明：自签名不会消除陌生外网用户看到的 SmartScreen 提示（与未签名相同）；' -ForegroundColor Yellow
    Write-Host '        它的价值是企业内导入 .cer 后可直接信任，以及为将来购买正式证书备好签名流程。' -ForegroundColor Yellow
}
