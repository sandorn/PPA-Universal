param(
    [ValidateSet("Register", "Unregister")]
    [string]$Action = "Register",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [string]$RegasmPath
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$addinDir = Join-Path $root "src\Hosts\PPA.Universal.ComAddIn\bin\$Configuration"
$universalDir = Join-Path $root "src\Hosts\PPA.Universal\bin\$Configuration"
$dllPath = Join-Path $addinDir "PPA.Universal.ComAddIn.dll"

if (-not (Test-Path $dllPath)) {
    throw "未找到 DLL：$dllPath，请先构建解决方案（dotnet build src\PPA.Layered.sln -c $Configuration）"
}

if (-not (Test-Path $universalDir)) {
    throw "未找到依赖目录：$universalDir，请先构建 PPA.Universal 项目"
}

# 复制 PPA.Universal 的依赖 DLL 至 COM Add-in 目录，确保 regasm 时可找到所有依赖
Write-Host "同步依赖 DLL 至 $addinDir"
Get-ChildItem $universalDir -Filter "*.dll" | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $addinDir $_.Name) -Force
}

if (-not $RegasmPath) {
    $frameworkDir = if ([Environment]::Is64BitOperatingSystem) {
        Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319"
    } else {
        Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319"
    }
    $RegasmPath = Join-Path $frameworkDir "regasm.exe"
}

if (-not (Test-Path $RegasmPath)) {
    throw "未找到 regasm，可通过 -RegasmPath 显式指定路径"
}

Write-Host "使用 regasm: $RegasmPath"
Write-Host "目标 DLL: $dllPath"

if ($Action -eq "Register") {
    & $RegasmPath $dllPath /codebase /verbose
    Write-Host "注册完成"
} else {
    & $RegasmPath $dllPath /unregister /verbose
    Write-Host "已撤销注册"
}

