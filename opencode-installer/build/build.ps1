# ============================================================
# build.ps1 - 构建 OpenCode 安装器 v4（便携版 / 厚重版）
# 用法：
#   powershell -ExecutionPolicy Bypass -File build\build.ps1
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -Edition portable
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -Sign
# ============================================================
[CmdletBinding()]
param(
    [ValidateSet('all', 'portable', 'full')]
    [string]$Edition = 'all',
    [switch]$SkipTests,
    [switch]$NoSmokeTest,
    [switch]$NoPackage,
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
$uiScript = Join-Path $srcDir 'lib\ui.ps1'
$version = '4.2.2.0'

. $coreScript

function Write-BuildStep { param([string]$Message) Write-Host "  [build] $Message" -ForegroundColor Cyan }

function Set-Utf8Bom {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($true)))
    Write-BuildStep "已补 UTF-8 BOM：$(Split-Path -Leaf $Path)"
}

function New-ZipFromDirectory {
    param([string]$Source, [string]$Output, [bool]$IncludeBaseDirectory)

    if (-not (Test-Path -LiteralPath $Source)) { throw "缺少打包源目录：$Source" }
    if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $Source,
        $Output,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $IncludeBaseDirectory
    )
    Write-BuildStep ("已打包 {0}：{1:N1} MB" -f (Split-Path -Leaf $Output), ((Get-Item $Output).Length / 1MB))
}

function Invoke-BuildTests {
    Write-BuildStep '运行单元测试...'
    $result = Invoke-Pester -Script (Join-Path $root 'tests\installer-core.Tests.ps1') -PassThru
    if (-not $result -or $result.FailedCount -gt 0) {
        throw "单元测试未通过（失败 $($result.FailedCount) 个），已中止构建"
    }
    Write-BuildStep "测试通过：$($result.PassedCount) 个"
}

function Test-EmbeddedResources {
    param([string]$ExePath, [string[]]$Names)

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

    $smokeRoot = Join-Path $env:TEMP ('opencode-installer-v4-smoke-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $smokeRoot | Out-Null
    try {
        $proc = Start-Process -FilePath $ExePath -ArgumentList @('-DryRun', '-TestRoot', $smokeRoot, '-NoPause') -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -ne 0) { throw "冒烟测试失败（退出码 $($proc.ExitCode)）" }
        Write-BuildStep '冒烟测试通过（-DryRun 全流程，退出码 0）'
    } finally {
        Remove-Item -LiteralPath $smokeRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-SelfTest {
    param([string]$ExePath)

    $proc = Start-Process -FilePath $ExePath -ArgumentList @('-SelfTest', '-NoPause') -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) { throw "自检失败（退出码 $($proc.ExitCode)）" }
    Write-BuildStep '自检通过（核心版本 / npm 命令 / 环境）'
}

function Publish-ReleaseEdition {
    param(
        [string]$EditionKey,
        [string]$EditionName,
        [string]$ExePath
    )

    $releaseRoot = Join-Path $PSScriptRoot 'release'
    $targetDir = Join-Path $releaseRoot ("opencode安装器-$EditionName")
    if (Test-Path -LiteralPath $targetDir) { Remove-Item -LiteralPath $targetDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    Copy-Item -LiteralPath $ExePath -Destination (Join-Path $targetDir 'opencode安装器.exe') -Force

    $guide = Join-Path $root ("docs\tutorials\opencode-installer-$EditionKey-guide.md")
    if (Test-Path -LiteralPath $guide) {
        Copy-Item -LiteralPath $guide -Destination (Join-Path $targetDir '使用教程.md') -Force
    } else {
        Write-BuildStep "缺少教程：$guide"
    }

    $pdf = Join-Path $PSScriptRoot ("tutorial-$EditionKey\tutorial.pdf")
    if (Test-Path -LiteralPath $pdf) {
        Copy-Item -LiteralPath $pdf -Destination (Join-Path $targetDir '使用教程.pdf') -Force
    } else {
        Write-BuildStep "缺少 PDF 教程：$pdf"
    }

    $zipPath = Join-Path $releaseRoot ("opencode安装器-$EditionName.zip")
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $targetDir,
        $zipPath,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $true
    )
    Write-BuildStep ("已打包 {0}：{1:N1} MB" -f (Split-Path -Leaf $zipPath), ((Get-Item $zipPath).Length / 1MB))
}

function Get-EditionList {
    param([string]$Edition)
    if ($Edition -eq 'all') { return @('portable', 'full') }
    return @($Edition)
}

# ============================================================
# 构建流程
# ============================================================

Write-Host ''
Write-Host '  ==== OpenCode 安装器 v4 构建 ====' -ForegroundColor Green
New-Item -ItemType Directory -Force -Path $distDir | Out-Null
foreach ($file in @($installScript, $coreScript, $uiScript)) { Set-Utf8Bom -Path $file }
if (-not $SkipTests) { Invoke-BuildTests } else { Write-BuildStep '跳过单元测试（-SkipTests）' }

$guizangZip = Join-Path $distDir 'guizang-ppt-skill.zip'
$outlookZip = Join-Path $distDir 'outlook-mcp.zip'
$excelZip = Join-Path $distDir 'excel-mcp.zip'
$pptZip = Join-Path $assetsDir 'ppt-master-skill-v6.7.3.zip'

New-ZipFromDirectory -Source (Join-Path $assetsDir 'guizang-ppt-skill') -Output $guizangZip -IncludeBaseDirectory $true
New-ZipFromDirectory -Source (Join-Path $assetsDir 'outlook-mcp') -Output $outlookZip -IncludeBaseDirectory $false
New-ZipFromDirectory -Source (Join-Path $assetsDir 'excel-mcp') -Output $excelZip -IncludeBaseDirectory $false

if (-not (Test-Path -LiteralPath $pptZip)) { throw "缺少 PPT-Master 6.7.3 打包：$pptZip" }
$zipSources = @{
    'ppt-master-skill-v6.7.3.zip' = $pptZip
    'guizang-ppt-skill.zip'       = $guizangZip
    'outlook-mcp.zip'             = $outlookZip
    'excel-mcp.zip'               = $excelZip
}

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

$editions = Get-EditionList -Edition $Edition
$cert = $null
if ($Sign) { $cert = Get-BuildCertificate -PfxPath $PfxPath -PfxPassword $PfxPassword }

foreach ($editionKey in $editions) {
    $info = Get-EditionInfo -Edition $editionKey
    $requiredPayloads = @(Get-RequiredPayloads -Edition $editionKey)
    $editionFile = Join-Path $distDir ("edition-{0}.txt" -f $editionKey)
    [System.IO.File]::WriteAllText($editionFile, $editionKey, (New-Object System.Text.UTF8Encoding($false)))

    $embedFiles = @{
        '%TEMP%\opencode-installer-v4\installer-core.ps1' = $coreScript
        '%TEMP%\opencode-installer-v4\ui.ps1'             = $uiScript
        '%TEMP%\opencode-installer-v4\edition.txt'        = $editionFile
    }
    foreach ($name in $requiredPayloads) {
        if (-not $zipSources.ContainsKey($name)) { throw "构建配置缺少资源：$name" }
        $embedFiles["%TEMP%\opencode-installer-v4\$name"] = $zipSources[$name]
    }

    $exePath = Join-Path $distDir $info.ExeName
    Write-BuildStep "编译 $($info.Name)（$($requiredPayloads.Count) 个资源包）..."
    if (Test-Path -LiteralPath $exePath) { Remove-Item -LiteralPath $exePath -Force }
    Invoke-PS2EXE -inputFile $installScript -outputFile $exePath `
        -title "OpenCode 安装器 - $($info.Name)" `
        -description "OpenCode 一键安装器：$($info.Description)" `
        -product 'OpenCode 安装器' `
        -version $version `
        -x64 `
        -embedFiles $embedFiles

    if (-not (Test-Path -LiteralPath $exePath)) { throw "编译失败：未生成 $($info.ExeName)" }
    Write-BuildStep ("exe 已生成：{0:N1} MB" -f ((Get-Item $exePath).Length / 1MB))

    $resourceNames = @('install.ps1', 'installer-core.ps1', 'ui.ps1', (Split-Path -Leaf $editionFile)) + $requiredPayloads
    Test-EmbeddedResources -ExePath $exePath -Names $resourceNames

    if (-not $NoSmokeTest) {
        Invoke-SmokeTest -ExePath $exePath
        Invoke-SelfTest -ExePath $exePath
    }
    if (-not $NoPackage) {
        Publish-ReleaseEdition -EditionKey $editionKey -EditionName $info.Name -ExePath $exePath
    }

    if ($Sign) {
        $signArgs = @{ FilePath = $exePath; Certificate = $cert; HashAlgorithm = 'SHA256' }
        if ($TimestampUrl) { $signArgs.TimestampServer = $TimestampUrl }
        $signature = Set-AuthenticodeSignature @signArgs
        if ($signature.SignerCertificate -and $signature.Status -eq 'UnknownError') {
            Write-BuildStep "签名已写入（发布者：$($signature.SignerCertificate.Subject)）；自签名根不受本机信任，状态 UnknownError 属预期"
        } else {
            Write-BuildStep "签名状态：$($signature.Status)"
        }
    }
}

if ($Sign -and $cert) {
    $cerPath = Join-Path $distDir 'OpenCode安装器-签名证书.cer'
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
    Write-BuildStep '已导出公钥证书（.cer，供企业内信任导入）'
}

Write-Host ''
Write-Host '  构建完成' -ForegroundColor Green
Get-ChildItem -LiteralPath $distDir -Filter 'OpenCode安装器-*.exe' | ForEach-Object {
    Write-Host ("  {0}  {1:N1} MB" -f $_.Name, ($_.Length / 1MB)) -ForegroundColor Cyan
}
if (-not $NoPackage) {
    Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'release') -Filter '*.zip' | ForEach-Object {
        Write-Host ("  {0}  {1:N1} MB" -f $_.Name, ($_.Length / 1MB)) -ForegroundColor Green
    }
}
if ($Sign) {
    Write-Host '  说明：自签名不会消除陌生外网用户看到的 SmartScreen 提示（与未签名相同）；' -ForegroundColor Yellow
    Write-Host '        它的价值是企业内导入 .cer 后可直接信任，以及为将来购买正式证书备好签名流程。' -ForegroundColor Yellow
}
