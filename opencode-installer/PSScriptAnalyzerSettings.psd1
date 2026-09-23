# PSScriptAnalyzer 项目配置
# 本项目是面向最终用户的交互式控制台安装器，以下规则为有意豁免（并非疏漏）：
@{
    ExcludeRules = @(
        # 交互式彩色控制台 UI 需要 Write-Host；Write-Output 无法实现配色/不换行等交互效果
        'PSAvoidUsingWriteHost',
        # 安装器自带 -DryRun 干跑模式，不需要 PowerShell 的 -WhatIf/ShouldProcess 机制
        'PSUseShouldProcessForStateChangingFunctions',
        # 这些函数名用复数更准确（返回集合 / 处理多个对象，如 Get-MissingPayloads）
        'PSUseSingularNouns'
    )
}
