# ui.ps1 - console UI helpers for OpenCode 安装器 v4

function Write-AppBanner {
    param([string]$Version, [hashtable]$EditionInfo)

    try { Clear-Host } catch { $null = $_ }
    Write-Host ''
    Write-Host '  ┌──────────────────────────────────────────────────────────┐' -ForegroundColor Cyan
    Write-Host "    OpenCode 一键安装器  v$Version · $($EditionInfo.Name)" -ForegroundColor Cyan
    Write-Host "    $($EditionInfo.Description)" -ForegroundColor DarkGray
    Write-Host '  └──────────────────────────────────────────────────────────┘' -ForegroundColor Cyan
    Write-Host ''
}

function Write-SectionTitle {
    param([string]$Title)

    Write-Host ''
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ('  ' + ('─' * 58)) -ForegroundColor DarkGray
}

function Write-KeyValue {
    param([string]$Label, [string]$Value, [string]$Status = '')

    $color = 'White'
    $icon = '·'
    switch ($Status) {
        'ok' { $color = 'Green'; $icon = '✓' }
        'warn' { $color = 'Yellow'; $icon = '!' }
        'error' { $color = 'Red'; $icon = '×' }
        default { $color = 'White'; $icon = '·' }
    }
    Write-Host ("  {0} {1} {2}" -f $Label.PadRight(10), $icon, $Value) -ForegroundColor $color
}

function Write-Status {
    param([string]$Status, [string]$Message, [string]$Detail = '')

    $icon = '·'
    $color = 'White'
    switch ($Status) {
        'ok' { $icon = '✓'; $color = 'Green' }
        'info' { $icon = 'i'; $color = 'Yellow' }
        'warn' { $icon = '!'; $color = 'DarkYellow' }
        'error' { $icon = '×'; $color = 'Red' }
        'plan' { $icon = '计划'; $color = 'Magenta' }
        'skip' { $icon = '-'; $color = 'DarkGray' }
        default { $icon = '·'; $color = 'White' }
    }
    $line = "  [$icon] $Message"
    if ($Detail) { $line += "  $Detail" }
    Write-Host $line -ForegroundColor $color
}

function Write-StepHeader {
    param([int]$Index, [int]$Total, [string]$Title)

    Write-Host ''
    Write-Host "  [$Index/$Total] $Title" -ForegroundColor Cyan
}

function Write-CompletionBanner {
    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════╗' -ForegroundColor Green
    Write-Host '  ║          OpenCode 安装完成！                  ║' -ForegroundColor Green
    Write-Host '  ╚══════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host ''
}

function Confirm-YesNo {
    param([string]$Question, [bool]$Default = $false)

    $suffix = if ($Default) { '(Y/n)' } else { '(y/N)' }
    while ($true) {
        $answer = Read-Host "  $Question $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        switch ($answer.Trim().ToLowerInvariant()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default { Write-Host '  请输入 y 或 n' -ForegroundColor Red }
        }
    }
}

function Read-SecretText {
    param([string]$Prompt)

    $secure = Read-Host "  $Prompt" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Wait-Key {
    if ($script:NoPause) { return }
    Write-Host ''
    Write-Host '  按任意键继续...' -ForegroundColor DarkGray
    try { $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { $null = Read-Host }
}
