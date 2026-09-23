$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$here\..\src\lib\installer-core.ps1"

$script:TmpRoot = Join-Path $env:TEMP ("opencode-installer-tests-" + [guid]::NewGuid().ToString('N'))

Describe 'Resolve-SkillAction' {
    It 'returns Install when nothing is installed' {
        Resolve-SkillAction -InstalledVersion $null -PayloadVersion '6.6.0' | Should Be 'Install'
    }
    It 'returns Update when payload is newer' {
        Resolve-SkillAction -InstalledVersion '6.5.0' -PayloadVersion '6.6.0' | Should Be 'Update'
    }
    It 'returns Skip when versions are equal' {
        Resolve-SkillAction -InstalledVersion '6.6.0' -PayloadVersion '6.6.0' | Should Be 'Skip'
    }
    It 'returns Update when installed version is unparseable' {
        Resolve-SkillAction -InstalledVersion 'weird' -PayloadVersion '6.6.0' | Should Be 'Update'
    }
}

Describe 'Read-InstallerManifest and Save-InstallerManifest' {
    It 'saves and reads back the manifest' {
        New-Item -ItemType Directory -Force -Path $script:TmpRoot | Out-Null
        $path = Join-Path $script:TmpRoot 'manifest.json'
        Save-InstallerManifest -Path $path -Manifest @{ installerVersion = '3.0.0'; skills = @{ 'ppt-master' = '6.6.0' } }
        $m = Read-InstallerManifest -Path $path
        $m.installerVersion | Should Be '3.0.0'
        $m.skills['ppt-master'] | Should Be '6.6.0'
    }
    It 'returns $null when the file is missing' {
        Read-InstallerManifest -Path (Join-Path $script:TmpRoot 'missing.json') | Should Be $null
    }
    It 'returns $null when the file is corrupt' {
        $path = Join-Path $script:TmpRoot 'corrupt.json'
        Set-Content -LiteralPath $path -Value '{ not json' -Encoding UTF8
        Read-InstallerManifest -Path $path | Should Be $null
    }
}

Describe 'Get-SkillVersionFromSkillMd' {
    It 'parses metadata version from front matter' {
        $p = Join-Path $script:TmpRoot 'SKILL.md'
        Set-Content -LiteralPath $p -Encoding UTF8 -Value @(
            '---'
            'name: ppt-master'
            'metadata:'
            '  version: "6.6.0"'
            '---'
            ''
            '# PPT Master'
        )
        Get-SkillVersionFromSkillMd -Path $p | Should Be '6.6.0'
    }
    It 'returns $null when front matter has no version' {
        $p = Join-Path $script:TmpRoot 'SKILL-no-version.md'
        Set-Content -LiteralPath $p -Encoding UTF8 -Value @('---', 'name: guizang-ppt-skill', '---')
        Get-SkillVersionFromSkillMd -Path $p | Should Be $null
    }
}

Describe 'Update-OpenCodeConfigJson' {
    It 'preserves existing keys and adds mcp and permission' {
        $json = '{"provider":{"deepseek":{"options":{"apiKey":"sk-test"}}},"model":"deepseek/deepseek-flash"}'
        $out = Update-OpenCodeConfigJson -ConfigJson $json -McpName 'outlook' -McpCommand @('python', 'outlook.py') -AllowSkillPermission $true | ConvertFrom-Json
        $out.provider.deepseek.options.apiKey | Should Be 'sk-test'
        $out.model | Should Be 'deepseek/deepseek-flash'
        $out.mcp.outlook.type | Should Be 'local'
        $out.mcp.outlook.command[0] | Should Be 'python'
        $out.permission.skill.'*' | Should Be 'allow'
    }
    It 'preserves other mcp servers' {
        $json = '{"mcp":{"my-excel":{"type":"local","command":["python","-m","excel_mcp","stdio"]}}}'
        $out = Update-OpenCodeConfigJson -ConfigJson $json -McpName 'outlook' -McpCommand @('python', 'outlook.py') -AllowSkillPermission $true | ConvertFrom-Json
        $out.mcp.'my-excel'.command[2] | Should Be 'excel_mcp'
        $out.mcp.outlook.type | Should Be 'local'
    }
    It 'is idempotent' {
        $once = Update-OpenCodeConfigJson -ConfigJson '{"mcp":{}}' -McpName 'outlook' -McpCommand @('python', 'outlook.py') -AllowSkillPermission $true
        $twice = Update-OpenCodeConfigJson -ConfigJson $once -McpName 'outlook' -McpCommand @('python', 'outlook.py') -AllowSkillPermission $true
        $twice | Should Be $once
    }
    It 'starts from an empty config when input is empty' {
        $out = Update-OpenCodeConfigJson -ConfigJson '' -McpName 'outlook' -McpCommand @('python', 'outlook.py') -AllowSkillPermission $true | ConvertFrom-Json
        $out.mcp.outlook.type | Should Be 'local'
    }
    It 'throws on invalid json' {
        { Update-OpenCodeConfigJson -ConfigJson '{ bad' -McpName 'outlook' -McpCommand @('python') -AllowSkillPermission $true } | Should Throw
    }
}

Describe 'Expand-SkillPack' {
    BeforeEach {
        $z = Join-Path $script:TmpRoot ('zp-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path "$z\src\ppt-master\skills\ppt-master\scripts" | Out-Null
        Set-Content -LiteralPath "$z\src\ppt-master\skills\ppt-master\SKILL.md" -Value 'x' -Encoding UTF8
        Set-Content -LiteralPath "$z\src\ppt-master\skills\ppt-master\scripts\a.py" -Value 'print(1)' -Encoding UTF8
        $script:zip = Join-Path $z 'pack.zip'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $z 'src'), $script:zip)
        $script:dest = Join-Path $z 'dest\ppt-master'
    }
    It 'extracts the inner root into the destination' {
        Expand-SkillPack -ZipPath $script:zip -InnerRoot 'ppt-master/skills/ppt-master' -Destination $script:dest
        Test-Path (Join-Path $script:dest 'SKILL.md') | Should Be $true
        Test-Path (Join-Path $script:dest 'scripts\a.py') | Should Be $true
        Test-Path (Join-Path $script:dest 'ppt-master') | Should Be $false
    }
    It 'replaces stale content but preserves .env' {
        New-Item -ItemType Directory -Force -Path $script:dest | Out-Null
        Set-Content -LiteralPath (Join-Path $script:dest 'stale.txt') -Value 'old' -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $script:dest '.env') -Value 'KEY=1' -Encoding UTF8
        Expand-SkillPack -ZipPath $script:zip -InnerRoot 'ppt-master/skills/ppt-master' -Destination $script:dest
        Test-Path (Join-Path $script:dest 'stale.txt') | Should Be $false
        (Get-Content -LiteralPath (Join-Path $script:dest '.env') -Raw).Trim() | Should Be 'KEY=1'
        Test-Path (Join-Path $script:dest 'SKILL.md') | Should Be $true
    }
}

Describe 'Get-PipInstallArguments' {
    It 'builds python -m pip install -r arguments with the mirror index' {
        $a = Get-PipInstallArguments -RequirementsPath 'C:\a\requirements.txt' -IndexUrl 'https://pypi.tuna.tsinghua.edu.cn/simple'
        $a[0] | Should Be '-m'
        $a[1] | Should Be 'pip'
        $a[2] | Should Be 'install'
        $a[3] | Should Be '-r'
        $a[4] | Should Be 'C:\a\requirements.txt'
        $a[5] | Should Be '-i'
        $a[6] | Should Be 'https://pypi.tuna.tsinghua.edu.cn/simple'
    }
}

Describe 'Get-MissingPayloads' {
    It 'returns only the payload names that are absent' {
        $root = Join-Path $script:TmpRoot ('pl-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $root | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'a.zip') -Value 'x' -Encoding UTF8
        $missing = @(Get-MissingPayloads -Root $root -Names @('a.zip', 'b.py'))
        $missing.Count | Should Be 1
        $missing[0] | Should Be 'b.py'
    }
    It 'returns an empty array when nothing is missing' {
        $root = Join-Path $script:TmpRoot ('pl2-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $root | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'a.zip') -Value 'x' -Encoding UTF8
        @(Get-MissingPayloads -Root $root -Names @('a.zip')).Count | Should Be 0
    }
}

Describe 'Add-SkillPermissionJson' {
    It 'adds permission.skill.* while preserving other keys' {
        $out = Add-SkillPermissionJson -ConfigJson '{"model":"x"}' | ConvertFrom-Json
        $out.model | Should Be 'x'
        $out.permission.skill.'*' | Should Be 'allow'
    }
    It 'is idempotent' {
        $once = Add-SkillPermissionJson -ConfigJson '{}'
        $twice = Add-SkillPermissionJson -ConfigJson $once
        $twice | Should Be $once
    }
}

Describe 'Merge-OpenCodeConfigJson' {
    It 'deep merges and lets the patch win' {
        $out = Merge-OpenCodeConfigJson -ConfigJson '{"provider":{"deepseek":{"options":{"apiKey":"sk-x"}}},"model":"old"}' -PatchJson '{"model":"new","provider":{"deepseek":{"name":"DeepSeek"}}}' | ConvertFrom-Json
        $out.model | Should Be 'new'
        $out.provider.deepseek.options.apiKey | Should Be 'sk-x'
        $out.provider.deepseek.name | Should Be 'DeepSeek'
    }
    It 'replaces arrays instead of appending' {
        $out = Merge-OpenCodeConfigJson -ConfigJson '{"a":[1,2]}' -PatchJson '{"a":[3]}' | ConvertFrom-Json
        $out.a.Count | Should Be 1
        $out.a[0] | Should Be 3
    }
    It 'throws on invalid patch json' {
        { Merge-OpenCodeConfigJson -ConfigJson '{}' -PatchJson '{ bad' } | Should Throw '不是合法 JSON'
    }
}

Describe 'Get-NodeDownloadCandidates' {
    It 'picks the newest LTS entry from the index' {
        $index = '[{"version":"v20.19.0","lts":"Iron"},{"version":"v22.19.0","lts":false},{"version":"v24.1.0","lts":"Krypton"}]'
        $c = @(Get-NodeDownloadCandidates -IndexJson $index)
        $c[0] | Should Be 'https://npmmirror.com/mirrors/node/v24.1.0/node-v24.1.0-x64.msi'
    }
    It 'falls back to pinned versions when the index cannot be parsed' {
        $c = @(Get-NodeDownloadCandidates -IndexJson 'not json')
        $c[0] | Should Be 'https://npmmirror.com/mirrors/node/v22.19.0/node-v22.19.0-x64.msi'
    }
}

Describe 'Get-PayloadRoot' {
    It 'prefers the temp root when every payload is present' {
        $root = Join-Path $script:TmpRoot ('pr-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $root | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'x.zip') -Value 'x' -Encoding UTF8
        Get-PayloadRoot -TempRoot $root -AssetsDir 'C:\assets' -RequiredNames @('x.zip') | Should Be $root
    }
    It 'falls back to the assets dir when the temp root is incomplete' {
        Get-PayloadRoot -TempRoot 'C:\definitely-missing' -AssetsDir 'C:\assets' -RequiredNames @('x.zip') | Should Be 'C:\assets'
    }
}

Describe 'Get-HostPortFromUrl' {
    It 'defaults to port 443 for https urls' {
        $endpoint = Get-HostPortFromUrl -Url 'https://api.deepseek.com/v1'
        $endpoint.Host | Should Be 'api.deepseek.com'
        $endpoint.Port | Should Be 443
    }
    It 'defaults to port 80 for http urls' {
        (Get-HostPortFromUrl -Url 'http://example.com/v1').Port | Should Be 80
    }
    It 'honours an explicit port' {
        (Get-HostPortFromUrl -Url 'http://localhost:11434/v1').Port | Should Be 11434
    }
    It 'returns $null for an unparseable url' {
        Get-HostPortFromUrl -Url 'not a url' | Should Be $null
    }
}

Describe 'Test-TcpEndpoint' {
    It 'returns true when the port is open' {
        $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
        $listener.Start()
        try {
            $port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
            Test-TcpEndpoint -HostName '127.0.0.1' -Port $port -TimeoutMs 2000 | Should Be $true
        } finally { $listener.Stop() }
    }
    It 'returns false when nothing listens on the port' {
        Test-TcpEndpoint -HostName '127.0.0.1' -Port 1 -TimeoutMs 2000 | Should Be $false
    }
}

Describe 'Resolve-ApiFailureMessage' {
    It 'explains an invalid key for status 401' {
        Resolve-ApiFailureMessage -StatusCode 401 -Endpoint 'api.deepseek.com:443' | Should Match '无效'
    }
    It 'explains rate limiting for status 429' {
        Resolve-ApiFailureMessage -StatusCode 429 -Endpoint 'api.deepseek.com:443' | Should Match '频繁'
    }
    It 'mentions the endpoint when the network failed' {
        Resolve-ApiFailureMessage -StatusCode 0 -Endpoint 'api.deepseek.com:443' | Should Match 'api\.deepseek\.com:443'
    }
}
