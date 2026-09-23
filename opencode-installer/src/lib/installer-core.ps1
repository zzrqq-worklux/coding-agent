# installer-core.ps1 - core logic for OpenCode 安装器 (v3)

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
