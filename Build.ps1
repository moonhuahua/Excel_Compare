param(
    [string]$Packager = (Join-Path $PSScriptRoot 'tools\Build-SingleFileApp.ps1'),
    [string]$OutputDir = (Join-Path $PSScriptRoot '发布'),
    [switch]$Test
)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$dotnetPath = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
if (-not (Test-Path -LiteralPath $dotnetPath)) {
    $dotnetPath = (Get-Command dotnet -ErrorAction Stop).Source
}
$referenceRoot = Join-Path $projectRoot 'src\packages\Microsoft.NETFramework.ReferenceAssemblies.net452.1.0.3\build'
if (-not (Test-Path -LiteralPath $referenceRoot)) { throw '缺少构建依赖，请先运行 restore_dependencies.py。' }
if (-not (Test-Path -LiteralPath $Packager)) { throw '缺少 tools\Build-SingleFileApp.ps1，也可使用 -Packager 指定打包脚本。' }
$workRoot = Join-Path $projectRoot 'build'
$releaseRoot = [IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $workRoot,$releaseRoot -Force | Out-Null
& $dotnetPath msbuild (Join-Path $projectRoot 'src\ExcelMerge.GUI\ExcelMerge.GUI.csproj') /nologo /v:minimal /p:Configuration=Release /p:TargetFrameworkVersion=v4.5.2 "/p:TargetFrameworkRootPath=$referenceRoot" '/p:NoWarn=1591' *> (Join-Path $workRoot '编译日志.txt')
if ($LASTEXITCODE -ne 0) { throw "编译失败，请检查 $workRoot\编译日志.txt" }
$binaryRoot = Join-Path $projectRoot 'src\ExcelMerge.GUI\bin\Release'
if ($Test) {
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
    & $compiler /nologo /target:exe "/out:$binaryRoot\ComparisonTests.exe" "/lib:$binaryRoot" /reference:ExcelMerge.dll,NPOI.dll,NPOI.OOXML.dll (Join-Path $projectRoot 'tests\ComparisonTests.cs')
    if ($LASTEXITCODE -ne 0) { throw '测试程序编译失败。' }
    & (Join-Path $binaryRoot 'ComparisonTests.exe') (Join-Path $projectRoot 'tests\fixtures') | Tee-Object -FilePath (Join-Path $workRoot '功能测试.txt')
    if ($LASTEXITCODE -ne 0) { throw '功能测试失败。' }
    $presentationRefs = Join-Path $referenceRoot '.NETFramework\v4.5.2'
    $presentationArgs = @('/reference:ExcelMerge.dll,ExcelMerge.GUI.exe,FastWpfGrid.dll,NPOI.dll,NPOI.OOXML.dll,Prism.Wpf.dll,Prism.dll')
    foreach ($reference in @('PresentationFramework.dll','PresentationCore.dll','WindowsBase.dll','System.Xaml.dll','Facades\System.Runtime.dll','Facades\System.ObjectModel.dll')) {
        $presentationArgs += ('/reference:' + (Join-Path $presentationRefs $reference))
    }
    & $compiler /nologo /target:exe "/out:$binaryRoot\PresentationTests.exe" "/lib:$binaryRoot" @presentationArgs (Join-Path $projectRoot 'tests\PresentationTests.cs')
    if ($LASTEXITCODE -ne 0) { throw '显示测试程序编译失败。' }
    & (Join-Path $binaryRoot 'PresentationTests.exe') (Join-Path $projectRoot 'tests\presentation-fixtures') | Tee-Object -FilePath (Join-Path $workRoot '显示功能测试.txt')
    if ($LASTEXITCODE -ne 0) { throw '显示功能测试失败。' }
}
$stageRoot = Join-Path $workRoot ('stage_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stageRoot | Out-Null
try {
    Get-ChildItem -LiteralPath $binaryRoot -File -Filter '*.dll' | Copy-Item -Destination $stageRoot
    Copy-Item -LiteralPath (Join-Path $binaryRoot 'ExcelMerge.GUI.exe'),(Join-Path $binaryRoot 'ExcelMerge.GUI.exe.config') -Destination $stageRoot
    foreach ($language in @('zh-CN','zh-Hans','ja-JP')) {
        $languageRoot = Join-Path $binaryRoot $language
        if (Test-Path -LiteralPath $languageRoot) { Copy-Item -LiteralPath $languageRoot -Destination $stageRoot -Recurse }
    }
    Copy-Item -LiteralPath (Join-Path $projectRoot 'src\README.md') -Destination (Join-Path $stageRoot 'ExcelMerge-LICENSE-and-README.md')
    $packageHost = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($packageHost) {
        & $packageHost.Source -NoProfile -ExecutionPolicy Bypass -File $Packager -SourceDir $stageRoot -EntryExe 'ExcelMerge.GUI.exe' -OutputDir $releaseRoot -OutputName 'Excel_Compare' -IconPath (Join-Path $projectRoot 'src\ExcelMerge.GUI\app.ico') -SilentMode TempSilent
        if ($LASTEXITCODE -ne 0) { throw '单文件封装失败。' }
    } else {
        # Windows PowerShell 5.1 needs a BOM to parse the packager's Chinese text.
        $compatiblePackager = Join-Path $workRoot 'packager_utf8bom.ps1'
        [IO.File]::WriteAllText($compatiblePackager, [IO.File]::ReadAllText($Packager, [Text.Encoding]::UTF8), [Text.UTF8Encoding]::new($true))
        try {
            & $compatiblePackager -SourceDir $stageRoot -EntryExe 'ExcelMerge.GUI.exe' -OutputDir $releaseRoot -OutputName 'Excel_Compare' -IconPath (Join-Path $projectRoot 'src\ExcelMerge.GUI\app.ico') -SilentMode TempSilent
        } finally { Remove-Item -LiteralPath $compatiblePackager -Force -ErrorAction SilentlyContinue }
    }
    if (-not (Test-Path -LiteralPath (Join-Path $releaseRoot 'Excel_Compare.exe'))) { throw '未生成单文件程序。' }
    Copy-Item -LiteralPath (Join-Path $projectRoot 'Excel选中两文件对比.ahk') -Destination $releaseRoot -Force
    Copy-Item -LiteralPath (Join-Path $projectRoot '使用说明.md') -Destination $releaseRoot -Force
} finally {
    $resolvedStage = [IO.Path]::GetFullPath($stageRoot)
    if ($resolvedStage.StartsWith([IO.Path]::GetFullPath($workRoot) + '\') -and (Test-Path -LiteralPath $resolvedStage)) {
        Remove-Item -LiteralPath $resolvedStage -Recurse -Force
    }
}
Get-Item -LiteralPath (Join-Path $releaseRoot 'Excel_Compare.exe') | Select-Object FullName,Length
