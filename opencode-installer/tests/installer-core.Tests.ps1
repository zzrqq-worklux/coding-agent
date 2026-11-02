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

Describe 'Get-EditionFromFile' {
    It 'reads portable from the first existing file' {
        $p = Join-Path $script:TmpRoot ('edition-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'portable' -Encoding ASCII
        Get-EditionFromFile -Paths @((Join-Path $script:TmpRoot 'missing.txt'), $p) | Should Be 'portable'
    }
    It 'reads full from file' {
        $p = Join-Path $script:TmpRoot ('edition-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'FULL' -Encoding ASCII
        Get-EditionFromFile -Paths @($p) | Should Be 'full'
    }
    It 'defaults to full when no file exists' {
        Get-EditionFromFile -Paths @((Join-Path $script:TmpRoot 'nope.txt')) | Should Be 'full'
    }
    It 'defaults to full for invalid content' {
        $p = Join-Path $script:TmpRoot ('edition-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'weird' -Encoding ASCII
        Get-EditionFromFile -Paths @($p) | Should Be 'full'
    }
}

Describe 'Get-EditionInfo' {
    It 'returns portable display information' {
        $i = Get-EditionInfo -Edition 'portable'
        $i.Key | Should Be 'portable'
        $i.Name | Should Be '便携版'
        $i.ExeName | Should Match '便携版'
    }
    It 'returns full display information' {
        $i = Get-EditionInfo -Edition 'full'
        $i.Key | Should Be 'full'
        $i.Name | Should Be '厚重版'
        $i.ExeName | Should Match '厚重版'
    }
    It 'falls back to full for an unknown edition' {
        (Get-EditionInfo -Edition 'weird').Key | Should Be 'full'
    }
}

Describe 'Get-EnabledComponents' {
    It 'portable enables only the two MCP servers' {
        (@(Get-EnabledComponents -Edition 'portable') -join ',') | Should Be 'outlook,excel'
    }
    It 'full enables skills and MCP servers' {
        (@(Get-EnabledComponents -Edition 'full') -join ',') | Should Be 'ppt-master,guizang,outlook,excel'
    }
    It 'unknown edition falls back to full' {
        (@(Get-EnabledComponents -Edition 'weird') -join ',') | Should Be 'ppt-master,guizang,outlook,excel'
    }
}

Describe 'Get-RequiredPayloads' {
    It 'portable excludes PPT and Guizang packages' {
        $p = @(Get-RequiredPayloads -Edition 'portable')
        ($p -contains 'outlook-mcp.zip') | Should Be $true
        ($p -contains 'excel-mcp.zip') | Should Be $true
        ($p -contains 'ppt-master-skill-v6.7.3.zip') | Should Be $false
        ($p -contains 'guizang-ppt-skill.zip') | Should Be $false
    }
    It 'full includes all four packages' {
        $p = @(Get-RequiredPayloads -Edition 'full')
        ($p -contains 'ppt-master-skill-v6.7.3.zip') | Should Be $true
        ($p -contains 'guizang-ppt-skill.zip') | Should Be $true
        ($p -contains 'outlook-mcp.zip') | Should Be $true
        ($p -contains 'excel-mcp.zip') | Should Be $true
    }
}

Describe 'Get-ConfigSectionStyle' {
    It 'detects a V1 provider section' {
        Get-ConfigSectionStyle -ConfigJson '{"provider":{"deepseek":{}}}' -Section 'provider' | Should Be 'V1'
    }
    It 'detects a V2 providers section' {
        Get-ConfigSectionStyle -ConfigJson '{"providers":{"deepseek":{}}}' -Section 'provider' | Should Be 'V2'
    }
    It 'defaults an empty provider section to V2' {
        Get-ConfigSectionStyle -ConfigJson '{}' -Section 'provider' | Should Be 'V2'
    }
    It 'detects a V1 direct MCP server section' {
        Get-ConfigSectionStyle -ConfigJson '{"mcp":{"outlook":{"type":"local"}}}' -Section 'mcp' | Should Be 'V1'
    }
    It 'detects a V2 mcp.servers section' {
        Get-ConfigSectionStyle -ConfigJson '{"mcp":{"servers":{}}}' -Section 'mcp' | Should Be 'V2'
    }
    It 'treats mcp.timeout alone as V2' {
        Get-ConfigSectionStyle -ConfigJson '{"mcp":{"timeout":{"startup":1}}}' -Section 'mcp' | Should Be 'V2'
    }
    It 'detects a V1 permission section' {
        Get-ConfigSectionStyle -ConfigJson '{"permission":{"skill":{"*":"allow"}}}' -Section 'permissions' | Should Be 'V1'
    }
    It 'detects a V2 permissions section' {
        Get-ConfigSectionStyle -ConfigJson '{"permissions":[]}' -Section 'permissions' | Should Be 'V2'
    }
    It 'throws on invalid JSON' {
        { Get-ConfigSectionStyle -ConfigJson '{ bad' -Section 'provider' } | Should Throw
    }
}

Describe 'ConvertTo-ModelDisplayName' {
    It 'prettifies a DeepSeek model ID' {
        ConvertTo-ModelDisplayName -ModelId 'deepseek-v4-pro' -ProviderName 'DeepSeek' | Should Be 'DeepSeek V4 Pro'
    }
    It 'prettifies the flash model' {
        ConvertTo-ModelDisplayName -ModelId 'deepseek-flash' -ProviderName 'DeepSeek' | Should Be 'DeepSeek Flash'
    }
    It 'keeps a custom provider display name' {
        ConvertTo-ModelDisplayName -ModelId 'qwen3-coder' -ProviderName 'SiliconFlow' | Should Be 'SiliconFlow Qwen3 Coder'
    }
}

Describe 'New-ProviderPatch' {
    It 'builds a native V2 provider patch' {
        $json = New-ProviderPatch -Style 'V2' -ProviderId 'deepseek' -Name 'DeepSeek' `
            -BaseUrl 'https://api.deepseek.com/v1' -ApiKeyRef '{file:C:/keys/deepseek}' `
            -Models @('deepseek-flash', 'deepseek-v4-pro') -MainModel 'deepseek-flash' -TitleModel 'deepseek-flash'
        $o = $json | ConvertFrom-Json
        $o.model | Should Be 'deepseek/deepseek-flash'
        $o.providers.deepseek.package | Should Be '@opencode/ai/providers/openai-compatible'
        $o.providers.deepseek.settings.baseURL | Should Be 'https://api.deepseek.com/v1'
        $o.providers.deepseek.settings.apiKey | Should Be '{file:C:/keys/deepseek}'
        $o.providers.deepseek.models.'deepseek-v4-pro'.name | Should Be 'DeepSeek V4 Pro'
        $o.agents.title.model | Should Be 'deepseek/deepseek-flash'
        ($o.PSObject.Properties.Name -contains 'provider') | Should Be $false
    }
    It 'builds a V1 provider patch' {
        $json = New-ProviderPatch -Style 'V1' -ProviderId 'deepseek' -Name 'DeepSeek' `
            -BaseUrl 'https://api.deepseek.com/v1' -ApiKeyRef '{file:C:/keys/deepseek}' `
            -Models @('deepseek-flash', 'deepseek-v4-pro') -MainModel 'deepseek-v4-pro' -TitleModel 'deepseek-flash'
        $o = $json | ConvertFrom-Json
        $o.provider.deepseek.npm | Should Be '@ai-sdk/openai-compatible'
        $o.provider.deepseek.options.baseURL | Should Be 'https://api.deepseek.com/v1'
        $o.provider.deepseek.options.apiKey | Should Be '{file:C:/keys/deepseek}'
        $o.small_model | Should Be 'deepseek/deepseek-flash'
        ($o.PSObject.Properties.Name -contains 'providers') | Should Be $false
    }
    It 'adds main and title models when they are missing from discovery' {
        $o = (New-ProviderPatch -Style 'V2' -ProviderId 'deepseek' -Name 'DeepSeek' -BaseUrl 'https://api.deepseek.com/v1' `
            -Models @('deepseek-flash') -MainModel 'deepseek-v4-pro' -TitleModel 'deepseek-flash') | ConvertFrom-Json
        ($o.providers.deepseek.models.PSObject.Properties.Name -contains 'deepseek-v4-pro') | Should Be $true
    }
    It 'omits apiKey when no reference is given' {
        $o = (New-ProviderPatch -Style 'V2' -ProviderId 'local' -Name 'Local' -BaseUrl 'http://127.0.0.1:1234/v1' `
            -Models @('coder') -MainModel 'coder' -TitleModel 'coder') | ConvertFrom-Json
        ($o.providers.local.settings.PSObject.Properties.Name -contains 'apiKey') | Should Be $false
    }
}

Describe 'New-McpServerPatch' {
    It 'builds a V2 mcp.servers patch' {
        $o = (New-McpServerPatch -Style 'V2' -McpName 'outlook' -Command @('python', '-m', 'outlook_com_mcp') `
            -Environment @{ OUTLOOK_MCP_DEFAULT_SIGNATURE = '对外' }) | ConvertFrom-Json
        $o.mcp.servers.outlook.type | Should Be 'local'
        $o.mcp.servers.outlook.command[2] | Should Be 'outlook_com_mcp'
        $o.mcp.servers.outlook.environment.OUTLOOK_MCP_DEFAULT_SIGNATURE | Should Be '对外'
        $o.mcp.timeout.startup | Should Be 60000
    }
    It 'preserves an existing V2 mcp timeout when merged' {
        $existing = '{"mcp":{"timeout":{"execution":600000},"servers":{"my-excel":{"type":"local","command":["python","-m","excel_mcp","stdio"]}}}}'
        $patch = New-McpServerPatch -Style 'V2' -McpName 'outlook' -Command @('python', '-m', 'outlook_com_mcp')
        $o = Merge-OpenCodeConfigJson -ConfigJson $existing -PatchJson $patch | ConvertFrom-Json
        $o.mcp.timeout.execution | Should Be 600000
        $o.mcp.timeout.startup | Should Be 60000
        $o.mcp.servers.'my-excel'.command[2] | Should Be 'excel_mcp'
    }
    It 'builds a V1 direct MCP patch' {
        $o = (New-McpServerPatch -Style 'V1' -McpName 'outlook' -Command @('python', 'server.py')) | ConvertFrom-Json
        $o.mcp.outlook.type | Should Be 'local'
        $o.mcp.outlook.command[1] | Should Be 'server.py'
    }
    It 'omits an empty environment object' {
        $o = (New-McpServerPatch -Style 'V2' -McpName 'outlook' -Command @('python', 'server.py')) | ConvertFrom-Json
        ($o.mcp.servers.outlook.PSObject.Properties.Name -contains 'environment') | Should Be $false
    }
}

Describe 'Add-SkillPermissionToJson' {
    It 'appends a V2 skill rule to an existing permissions array' {
        $o = (Add-SkillPermissionToJson -ConfigJson '{"permissions":[{"action":"read","resource":"*","effect":"allow"}]}' -Style 'V2') | ConvertFrom-Json
        $o.permissions.Count | Should Be 2
        $o.permissions[0].action | Should Be 'read'
        $o.permissions[1].action | Should Be 'skill'
        $o.permissions[1].resource | Should Be '*'
        $o.permissions[1].effect | Should Be 'allow'
    }
    It 'is idempotent for V2' {
        $once = Add-SkillPermissionToJson -ConfigJson '{}' -Style 'V2'
        $twice = Add-SkillPermissionToJson -ConfigJson $once -Style 'V2'
        $twice | Should Be $once
    }
    It 'adds a V1 permission and preserves other permission keys' {
        $o = (Add-SkillPermissionToJson -ConfigJson '{"permission":{"bash":{"git push *":"ask"}}}' -Style 'V1') | ConvertFrom-Json
        $o.permission.bash.'git push *' | Should Be 'ask'
        $o.permission.skill.'*' | Should Be 'allow'
    }
    It 'throws on invalid JSON' {
        { Add-SkillPermissionToJson -ConfigJson '{ bad' -Style 'V2' } | Should Throw
    }
}

Describe 'Resolve-ConfigValueReference' {
    It 'reads a file reference' {
        $p = Join-Path $script:TmpRoot ('secret-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'sk-test' -Encoding ASCII -NoNewline
        Resolve-ConfigValueReference -Value "{file:$p}" -ConfigDir $script:TmpRoot | Should Be 'sk-test'
    }
    It 'resolves a relative file reference against the config dir' {
        $name = 'secret-' + [guid]::NewGuid().ToString('N') + '.txt'
        $p = Join-Path $script:TmpRoot $name
        Set-Content -LiteralPath $p -Value 'sk-relative' -Encoding ASCII -NoNewline
        Resolve-ConfigValueReference -Value "{file:$name}" -ConfigDir $script:TmpRoot | Should Be 'sk-relative'
    }
    It 'reads an environment reference' {
        $env:OPENCODE_INSTALLER_TEST_KEY = 'sk-env'
        try {
            Resolve-ConfigValueReference -Value '{env:OPENCODE_INSTALLER_TEST_KEY}' -ConfigDir $script:TmpRoot | Should Be 'sk-env'
        } finally {
            Remove-Item Env:\OPENCODE_INSTALLER_TEST_KEY -ErrorAction SilentlyContinue
        }
    }
    It 'returns a literal value unchanged' {
        Resolve-ConfigValueReference -Value 'sk-literal' -ConfigDir $script:TmpRoot | Should Be 'sk-literal'
    }
    It 'returns null for a missing file' {
        Resolve-ConfigValueReference -Value '{file:missing.txt}' -ConfigDir $script:TmpRoot | Should Be $null
    }
}

Describe 'Get-DefaultComponentSelection' {
    It 'portable selects only Outlook and Excel by default' {
        $s = @(Get-DefaultComponentSelection -Edition 'portable')
        $s.Count | Should Be 2
        $s[0].Key | Should Be 'outlook'
        $s[1].Key | Should Be 'excel'
        $s[0].Selected | Should Be $true
    }
    It 'full selects all four by default' {
        $s = @(Get-DefaultComponentSelection -Edition 'full')
        $s.Count | Should Be 4
        ($s | Where-Object { $_.Key -eq 'ppt-master' }).Selected | Should Be $true
    }
    It 'reports whether a component is enabled for an edition' {
        Test-ComponentEnabled -Edition 'portable' -Key 'ppt-master' | Should Be $false
        Test-ComponentEnabled -Edition 'full' -Key 'ppt-master' | Should Be $true
    }
}

Describe 'Get-EffectivePermissionStyle' {
    It 'uses V1 when the provider section is V1 and no permissions exist' {
        Get-EffectivePermissionStyle -ConfigJson '{"provider":{"deepseek":{}}}' | Should Be 'V1'
    }
    It 'uses V2 for a V2 provider config' {
        Get-EffectivePermissionStyle -ConfigJson '{"providers":{"deepseek":{}}}' | Should Be 'V2'
    }
    It 'keeps V2 permissions even with a V1 provider' {
        Get-EffectivePermissionStyle -ConfigJson '{"provider":{},"permissions":[]}' | Should Be 'V2'
    }
    It 'defaults to V2 for an empty config' {
        Get-EffectivePermissionStyle -ConfigJson '' | Should Be 'V2'
    }
}

Describe 'Get-ApiModelsUrl' {
    It 'appends models to a base URL without a trailing slash' {
        Get-ApiModelsUrl -BaseUrl 'https://api.deepseek.com/v1' | Should Be 'https://api.deepseek.com/v1/models'
    }
    It 'handles a trailing slash' {
        Get-ApiModelsUrl -BaseUrl 'https://api.deepseek.com/v1/' | Should Be 'https://api.deepseek.com/v1/models'
    }
    It 'returns null for an empty base URL' {
        Get-ApiModelsUrl -BaseUrl '' | Should Be $null
    }
}

Describe 'ConvertFrom-ModelsApiResponse' {
    It 'parses an OpenAI-compatible data array' {
        $m = @(ConvertFrom-ModelsApiResponse -Json '{"object":"list","data":[{"id":"deepseek-flash"},{"id":"deepseek-v4-pro"}]}')
        ($m -join ',') | Should Be 'deepseek-flash,deepseek-v4-pro'
    }
    It 'parses a models array' {
        $m = @(ConvertFrom-ModelsApiResponse -Json '{"models":[{"id":"a"},{"id":"b"}]}')
        ($m -join ',') | Should Be 'a,b'
    }
    It 'parses a plain string array' {
        $m = @(ConvertFrom-ModelsApiResponse -Json '["a","b"]')
        ($m -join ',') | Should Be 'a,b'
    }
    It 'deduplicates and filters empty IDs' {
        $m = @(ConvertFrom-ModelsApiResponse -Json '{"data":[{"id":"a"},{"id":"a"},{"id":""},{"id":null}]}')
        $m.Count | Should Be 1
        $m[0] | Should Be 'a'
    }
    It 'returns an empty array for invalid JSON' {
        @(ConvertFrom-ModelsApiResponse -Json '{ bad').Count | Should Be 0
    }
    It 'returns an empty array for empty input' {
        @(ConvertFrom-ModelsApiResponse -Json '').Count | Should Be 0
    }
}

Describe 'Select-DefaultModel' {
    It 'prefers deepseek-flash' {
        Select-DefaultModel -Models @('deepseek-v4-pro', 'deepseek-flash') | Should Be 'deepseek-flash'
    }
    It 'prefers any model containing flash' {
        Select-DefaultModel -Models @('foo', 'bar-flash-1') | Should Be 'bar-flash-1'
    }
    It 'falls back to the first model' {
        Select-DefaultModel -Models @('a', 'b') | Should Be 'a'
    }
    It 'returns null for an empty list' {
        Select-DefaultModel -Models @() | Should Be $null
    }
}

Describe 'Get-ApiKeyReference' {
    It 'normalizes backslashes to forward slashes' {
        Get-ApiKeyReference -Path 'C:\Users\me\.secrets\deepseek-api-key' | Should Be '{file:C:/Users/me/.secrets/deepseek-api-key}'
    }
    It 'keeps forward slashes' {
        Get-ApiKeyReference -Path 'C:/keys/deepseek' | Should Be '{file:C:/keys/deepseek}'
    }
    It 'returns null for an empty path' {
        Get-ApiKeyReference -Path '' | Should Be $null
    }
}

Describe 'Get-MajorVersion' {
    It 'parses a V1 version' {
        Get-MajorVersion -Version '1.19.5' | Should Be 1
    }
    It 'parses a V2 version' {
        Get-MajorVersion -Version '2.2.2' | Should Be 2
    }
    It 'parses a leading v' {
        Get-MajorVersion -Version 'v2.0.0' | Should Be 2
    }
    It 'returns null for garbage' {
        Get-MajorVersion -Version 'garbage' | Should Be $null
    }
    It 'returns null for an empty version' {
        Get-MajorVersion -Version '' | Should Be $null
    }
}

Describe 'Get-OpenCodeInstallMethod' {
    It 'detects an npm-managed install' {
        Get-OpenCodeInstallMethod -Path 'C:\Users\me\AppData\Roaming\npm\opencode.cmd' -NpmPrefix 'C:\Users\me\AppData\Roaming\npm' | Should Be 'npm'
    }
    It 'detects a curl-managed install' {
        Get-OpenCodeInstallMethod -Path 'C:\Users\me\.opencode\bin\opencode.exe' -NpmPrefix 'C:\Users\me\AppData\Roaming\npm' | Should Be 'curl'
    }
    It 'returns unknown for a missing path' {
        Get-OpenCodeInstallMethod -Path '' -NpmPrefix 'C:\Users\me\AppData\Roaming\npm' | Should Be 'unknown'
    }
    It 'returns unknown without an npm prefix' {
        Get-OpenCodeInstallMethod -Path 'C:\Users\me\AppData\Roaming\npm\opencode.cmd' -NpmPrefix '' | Should Be 'unknown'
    }
}

Describe 'Get-OpenCodeInstallPlan' {
    It 'installs @opencode/cli when OpenCode is absent' {
        $p = Get-OpenCodeInstallPlan -State @{ Version = ''; Path = ''; NpmPrefix = 'C:\npm'; HasV1 = $false; HasV2 = $false }
        $p.Action | Should Be 'Install'
        $p.Commands[0][0] | Should Be 'npm'
        $p.Commands[0][1] | Should Be 'install'
        $p.Commands[0][3] | Should Be '@opencode/cli@latest'
    }
    It 'migrates a V1 npm install to V2' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = '1.19.5'; Path = 'C:\npm\opencode.cmd'; NpmPrefix = 'C:\npm'; HasV1 = $true; HasV2 = $false
        }
        $p.Action | Should Be 'MigrateV1'
        $p.Commands[0][1] | Should Be 'uninstall'
        $p.Commands[0][3] | Should Be 'opencode-ai'
        $p.Commands[1][3] | Should Be '@opencode/cli@latest'
    }
    It 'updates an npm-managed V2 install' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = '2.2.2'; Path = 'C:\npm\opencode.cmd'; NpmPrefix = 'C:\npm'; HasV1 = $false; HasV2 = $true
        }
        $p.Action | Should Be 'UpdateNpm'
        $p.Commands[0][0] | Should Be 'opencode'
        $p.Commands[0][1] | Should Be 'upgrade'
        $p.Commands[0][3] | Should Be 'npm'
    }
    It 'reinstalls when the V2 package exists but the opencode command is missing' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = ''; Path = ''; NpmPrefix = 'C:\npm'; HasV1 = $false; HasV2 = $true
        }
        $p.Action | Should Be 'Install'
        $p.Commands[0][1] | Should Be 'install'
        $p.Commands[0][3] | Should Be '@opencode/cli@latest'
    }
    It 'migrates when both V1 and V2 packages are present' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = '2.2.2'; Path = 'C:\npm\opencode.cmd'; NpmPrefix = 'C:\npm'; HasV1 = $true; HasV2 = $true
        }
        $p.Action | Should Be 'MigrateV1'
    }
    It 'updates a curl-managed V2 install' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = '2.2.2'; Path = 'C:\Users\me\.opencode\bin\opencode.exe'; NpmPrefix = 'C:\npm'; HasV1 = $false; HasV2 = $true
        }
        $p.Action | Should Be 'UpdateCurl'
        $p.Commands[0][0] | Should Be 'opencode'
        $p.Commands[0][1] | Should Be 'upgrade'
    }
    It 'returns None when already at the latest version' {
        $p = Get-OpenCodeInstallPlan -State @{
            Version = '2.2.2'; Path = 'C:\npm\opencode.cmd'; NpmPrefix = 'C:\npm'; HasV1 = $false; HasV2 = $true; LatestVersion = '2.2.2'
        }
        $p.Action | Should Be 'None'
        $p.Commands.Count | Should Be 0
    }
}

Describe 'Get-NpmCommand' {
    It 'prefers npm.cmd to avoid the npm.ps1 shim' {
        Get-NpmCommand -Lookup { param($name) if ($name -eq 'npm.cmd') { 'npm.cmd' } } | Should Be 'npm.cmd'
    }
    It 'falls back to npm when npm.cmd is unavailable' {
        Get-NpmCommand -Lookup { param($name) if ($name -eq 'npm') { 'npm' } } | Should Be 'npm'
    }
    It 'falls back to npm when nothing is found' {
        Get-NpmCommand -Lookup { param($name) $null } | Should Be 'npm'
    }
    It 'never calls bare npm from install.ps1' {
        $text = Get-Content (Join-Path $here '..\src\install.ps1') -Raw
        ([regex]::Matches($text, '&\s+npm\s')).Count | Should Be 0
    }
}

Describe 'Resolve-NpmExecutable' {
    It 'prefers npm.cmd and returns a full path' {
        $cmd = Get-Command cmd.exe
        (Resolve-NpmExecutable -Lookup { param($name) if ($name -eq 'npm.cmd') { $cmd } }) | Should Be $cmd.Source
    }
    It 'accepts string lookups for testability' {
        Resolve-NpmExecutable -Lookup { param($name) if ($name -eq 'npm.cmd') { 'npm.cmd' } } | Should Be 'npm.cmd'
    }
    It 'rejects the npm.ps1 shim' {
        $shimPath = Join-Path $script:TmpRoot ('npm-shim-' + [guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $shimPath -Value '$null' -Encoding UTF8
        $shim = Get-Command $shimPath
        (Resolve-NpmExecutable -Lookup { param($name) if ($name -eq 'npm') { $shim } }) | Should Be $null
    }
    It 'returns null when nothing is found' {
        (Resolve-NpmExecutable -Lookup { param($name) $null }) | Should Be $null
    }
}

Describe 'Test-SafeNpmCommand' {
    It 'rejects an npm.ps1 shim' {
        $shimPath = Join-Path $script:TmpRoot ('npm-unsafe-' + [guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $shimPath -Value '$null' -Encoding UTF8
        (Test-SafeNpmCommand -Command $shimPath) | Should Be $false
    }
    It 'accepts an application command' {
        (Test-SafeNpmCommand -Command (Get-Command cmd.exe).Source) | Should Be $true
    }
    It 'rejects an empty command' {
        (Test-SafeNpmCommand -Command '') | Should Be $false
    }
}

Describe 'Get-ProxyConfiguration' {
    It 'reads HTTPS_PROXY first' {
        $oldHttps = $env:HTTPS_PROXY
        $oldHttp = $env:HTTP_PROXY
        try {
            $env:HTTPS_PROXY = 'http://proxy.local:8080'
            $env:HTTP_PROXY = 'http://fallback.local:8080'
            $proxy = Get-ProxyConfiguration -Uri 'https://example.com'
            $proxy.Uri | Should Be 'http://proxy.local:8080'
            $proxy.Source | Should Be 'HTTPS_PROXY'
        } finally {
            $env:HTTPS_PROXY = $oldHttps
            $env:HTTP_PROXY = $oldHttp
        }
    }
    It 'falls back to HTTP_PROXY' {
        $oldHttps = $env:HTTPS_PROXY
        $oldHttp = $env:HTTP_PROXY
        try {
            $env:HTTPS_PROXY = $null
            $env:HTTP_PROXY = 'http://fallback.local:8080'
            $proxy = Get-ProxyConfiguration -Uri 'https://example.com'
            $proxy.Uri | Should Be 'http://fallback.local:8080'
            $proxy.Source | Should Be 'HTTP_PROXY'
        } finally {
            $env:HTTPS_PROXY = $oldHttps
            $env:HTTP_PROXY = $oldHttp
        }
    }
}

Describe 'Test-GitBashEnvironment' {
    It 'detects MSYSTEM' {
        $old = $env:MSYSTEM
        try {
            $env:MSYSTEM = 'MINGW64'
            (Test-GitBashEnvironment) | Should Be $true
        } finally { $env:MSYSTEM = $old }
    }
    It 'detects OSTYPE msys' {
        $oldSystem = $env:MSYSTEM
        $oldType = $env:OSTYPE
        try {
            $env:MSYSTEM = $null
            $env:OSTYPE = 'msys'
            (Test-GitBashEnvironment) | Should Be $true
        } finally {
            $env:MSYSTEM = $oldSystem
            $env:OSTYPE = $oldType
        }
    }
    It 'detects a bash SHELL' {
        $oldSystem = $env:MSYSTEM
        $oldType = $env:OSTYPE
        $oldShell = $env:SHELL
        try {
            $env:MSYSTEM = $null
            $env:OSTYPE = $null
            $env:SHELL = '/usr/bin/bash'
            (Test-GitBashEnvironment) | Should Be $true
        } finally {
            $env:MSYSTEM = $oldSystem
            $env:OSTYPE = $oldType
            $env:SHELL = $oldShell
        }
    }
    It 'returns false in a Windows PowerShell session' {
        $oldSystem = $env:MSYSTEM
        $oldType = $env:OSTYPE
        $oldShell = $env:SHELL
        try {
            $env:MSYSTEM = $null
            $env:OSTYPE = $null
            $env:SHELL = $null
            (Test-GitBashEnvironment) | Should Be $false
        } finally {
            $env:MSYSTEM = $oldSystem
            $env:OSTYPE = $oldType
            $env:SHELL = $oldShell
        }
    }
}

Describe 'Get-PowerShell7InstallDir' {
    It 'uses the per-user Programs directory' {
        $dir = Get-PowerShell7InstallDir
        ($dir -like '*\Programs\PowerShell\7') | Should Be $true
        ($dir -like '*Program Files*') | Should Be $false
    }
}

Describe 'PowerShell 7 pinned release' {
    It 'pins the official 7.6.9 x64 zip hash' {
        $script:PowerShell7Version | Should Be '7.6.9'
        $script:PowerShell7Sha256 | Should Be 'CE7484804C0F0AD4DDA10C958E71ACDAF7EEBFB22505C1B480F89E0E9F1DA94D'
    }
}

Describe 'Get-PowerShell7DownloadUrls' {
    It 'includes GitHub and the TUNA mirror' {
        $old = $env:OC_POWERSHELL_URL
        try {
            $env:OC_POWERSHELL_URL = $null
            $urls = @(Get-PowerShell7DownloadUrls)
            ($urls -join "`n") | Should Match 'github.com/PowerShell/PowerShell/releases'
            ($urls -join "`n") | Should Match 'mirrors.tuna.tsinghua.edu.cn'
        } finally { $env:OC_POWERSHELL_URL = $old }
    }
    It 'puts a custom mirror first' {
        $old = $env:OC_POWERSHELL_URL
        try {
            $env:OC_POWERSHELL_URL = 'https://mirror.local/pwsh.zip'
            $urls = @(Get-PowerShell7DownloadUrls)
            $urls[0] | Should Be 'https://mirror.local/pwsh.zip'
        } finally { $env:OC_POWERSHELL_URL = $old }
    }
}

Describe 'Test-FileSha256' {
    It 'accepts the matching hash' {
        $p = Join-Path $script:TmpRoot ('hash-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'opencode' -Encoding UTF8
        $hash = (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash
        (Test-FileSha256 -Path $p -Expected $hash) | Should Be $true
    }
    It 'rejects a mismatching hash' {
        $p = Join-Path $script:TmpRoot ('hash-bad-' + [guid]::NewGuid().ToString('N') + '.txt')
        Set-Content -LiteralPath $p -Value 'opencode' -Encoding UTF8
        (Test-FileSha256 -Path $p -Expected '0000000000000000000000000000000000000000000000000000000000000000') | Should Be $false
    }
    It 'returns false for a missing file' {
        (Test-FileSha256 -Path (Join-Path $script:TmpRoot 'missing-hash.zip') -Expected 'x') | Should Be $false
    }
}

Describe 'Get-PowerShell7LocalZipCandidates' {
    It 'finds the pinned PowerShell zip in a directory' {
        $dir = Join-Path $script:TmpRoot ('pwsh-local-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $zip = Join-Path $dir ("PowerShell-$($script:PowerShell7Version)-win-x64.zip")
        Set-Content -LiteralPath $zip -Value 'zip' -Encoding UTF8
        $found = @(Get-PowerShell7LocalZipCandidates -Directories @($dir))
        ($found -contains $zip) | Should Be $true
    }
}

Describe 'Get-PowerShell7Version' {
    It 'parses the version from an explicit pwsh path' {
        $fake = Join-Path $script:TmpRoot ('fake-pwsh-' + [guid]::NewGuid().ToString('N') + '.cmd')
        Set-Content -LiteralPath $fake -Value @('@echo off', 'echo 7.6.9') -Encoding ASCII
        (Get-PowerShell7Version -Path $fake) | Should Be '7.6.9'
    }
    It 'returns null for a missing explicit path' {
        (Get-PowerShell7Version -Path (Join-Path $script:TmpRoot 'missing-pwsh.exe')) | Should Be $null
    }
}

Describe 'Get-OpenCodeLauncherContent' {
    It 'supports dragged folders and a pwsh fallback' {
        $text = Get-OpenCodeLauncherContent
        $text | Should Match '%~1'
        $text | Should Match 'pwsh'
        $text | Should Match 'opencode'
        $text | Should Match 'cd /d'
    }
}

Describe 'New-OpenCodeShortcut' {
    It 'creates the launcher and desktop shortcut' {
        $desktop = Join-Path $script:TmpRoot ('desktop-' + [guid]::NewGuid().ToString('N'))
        $launcherDir = Join-Path $script:TmpRoot ('launcher-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path $desktop | Out-Null
        $lnk = New-OpenCodeShortcut -DesktopPath $desktop -LauncherDir $launcherDir
        (Test-Path -LiteralPath $lnk) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $launcherDir 'OpenCode.cmd')) | Should Be $true
    }
}

Describe 'Expand-ZipPackage' {
    It 'extracts a package zip into the destination' {
        $root = Join-Path $script:TmpRoot ('pkg-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path "$root\src\pkg" | Out-Null
        Set-Content -LiteralPath "$root\src\pkg\__init__.py" -Value 'x' -Encoding UTF8
        $zip = Join-Path $root 'pkg.zip'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $root 'src'), $zip)
        $dest = Join-Path $root 'out\pkg'
        Expand-ZipPackage -ZipPath $zip -Destination $dest
        Test-Path (Join-Path $dest 'pkg\__init__.py') | Should Be $true
    }
    It 'removes stale content before extracting' {
        $root = Join-Path $script:TmpRoot ('pkg2-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Force -Path "$root\src\pkg" | Out-Null
        Set-Content -LiteralPath "$root\src\pkg\a.txt" -Value 'new' -Encoding UTF8
        $zip = Join-Path $root 'pkg.zip'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $root 'src'), $zip)
        $dest = Join-Path $root 'out'
        New-Item -ItemType Directory -Force -Path $dest | Out-Null
        Set-Content -LiteralPath (Join-Path $dest 'stale.txt') -Value 'old' -Encoding UTF8
        Expand-ZipPackage -ZipPath $zip -Destination $dest
        Test-Path (Join-Path $dest 'stale.txt') | Should Be $false
        Test-Path (Join-Path $dest 'pkg\a.txt') | Should Be $true
    }
    It 'throws when the zip is missing' {
        { Expand-ZipPackage -ZipPath (Join-Path $script:TmpRoot 'missing.zip') -Destination (Join-Path $script:TmpRoot 'x') } | Should Throw
    }
}

Describe 'Get-PackageVersionFromPyProject' {
    It 'parses the project version' {
        $p = Join-Path $script:TmpRoot ('pyproject-' + [guid]::NewGuid().ToString('N') + '.toml')
        Set-Content -LiteralPath $p -Encoding UTF8 -Value @(
            '[build-system]'
            'requires = ["setuptools"]'
            ''
            '[project]'
            'name = "outlook-com-mcp"'
            'version = "0.1.0"'
        )
        Get-PackageVersionFromPyProject -Path $p | Should Be '0.1.0'
    }
    It 'parses a post-release version' {
        $p = Join-Path $script:TmpRoot ('pyproject-' + [guid]::NewGuid().ToString('N') + '.toml')
        Set-Content -LiteralPath $p -Encoding UTF8 -Value @('[project]', 'version = "0.1.8.post1"')
        Get-PackageVersionFromPyProject -Path $p | Should Be '0.1.8.post1'
    }
    It 'returns null when the file is missing' {
        Get-PackageVersionFromPyProject -Path (Join-Path $script:TmpRoot 'missing.toml') | Should Be $null
    }
    It 'returns null when there is no project version' {
        $p = Join-Path $script:TmpRoot ('pyproject-' + [guid]::NewGuid().ToString('N') + '.toml')
        Set-Content -LiteralPath $p -Encoding UTF8 -Value @('[build-system]', 'requires = []')
        Get-PackageVersionFromPyProject -Path $p | Should Be $null
    }
}
