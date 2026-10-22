<#
.SYNOPSIS
    通用 7z SFX 单文件版软件打包制作工具 (Single-File Portable App Packager)

.DESCRIPTION
    将绿色版/便携版软件文件夹打包为单个独立的 .exe 可执行文件。
    支持功能：
    1. 自动解压到临时目录静默运行，主程序退出后自动清理临时文件，不留磁盘垃圾；
    2. 自动从源主程序中提取高清多分辨率内嵌图标 (.ico) 并无损注入到生成的单文件 EXE 中；
    3. 支持预打包清理瘦身（保守移除无用多国语言包、pdb 调试符号，自动备份）；
    4. 自动依赖管理：若本地未安装 7-Zip/LZMA SDK，自动下载并缓存至工具目录；
    5. 全程支持中文路径与 UTF-8 编码。

.PARAMETER SourceDir
    必填。需要打包的绿色软件/程序文件夹路径。

.PARAMETER EntryExe
    必填。文件夹内的主可执行文件名或相对路径（如 "App.exe" 或 "bin\App.exe"）。

.PARAMETER OutputName
    可选。输出的单文件 EXE 名称（不含 .exe 扩展名）。默认使用源文件夹名称或主程序名。

.PARAMETER OutputDir
    可选。输出目录。默认输出到源文件夹的上级目录。

.PARAMETER IconPath
    可选。自定义 ICO 图标文件路径。若未指定，则自动从 EntryExe 中提取原始图标。

.PARAMETER Clean
    开关。打包前对源文件夹进行安全瘦身（移除冗余多语言包与 pdb 文件，保留备份）。

.PARAMETER KeepArchive
    开关。是否保留中间生成的 .7z 压缩包。默认自动清理，只保留最终单文件 .exe。

.PARAMETER CompressionLevel
    可选。压缩级别（1-9）。默认为 9 (最大极限压缩)。

.PARAMETER SilentMode
    可选。运行模式：
    - "TempSilent" (默认): 静默解压至 %TEMP%，运行主程序，退出后自动清理。
    - "TempShowProgress": 显示解压进度条后运行，退出清理。
    - "CurrentDir": 解压至当前运行目录同级并运行（便携解压模式）。

.EXAMPLE
    pwsh -ExecutionPolicy Bypass -File .\scripts\Build-SingleFileApp.ps1 `
        -SourceDir "C:\Tools\MyTool" `
        -EntryExe "MyTool.exe" `
        -Clean
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SourceDir,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$EntryExe,

    [Parameter(Mandatory = $false)]
    [string]$OutputName,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [string]$IconPath,

    [Parameter(Mandatory = $false)]
    [switch]$Clean,

    [Parameter(Mandatory = $false)]
    [switch]$KeepArchive,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 9)]
    [int]$CompressionLevel = 9,

    [Parameter(Mandatory = $false)]
    [ValidateSet("TempSilent", "TempShowProgress", "CurrentDir")]
    [string]$SilentMode = "TempSilent",

    [Parameter(Mandatory = $false)]
    [string[]]$KeepCultures = @('zh-CN', 'zh-Hans', 'zh-Hant', 'zh-TW', 'en-US', 'en-GB', 'en')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Log([string]$Message, [string]$Level = "INFO") {
    $time = Get-Date -Format 'HH:mm:ss'
    switch ($Level) {
        "INFO"    { Write-Host "[$time] [信息] $Message" -ForegroundColor Cyan }
        "SUCCESS" { Write-Host "[$time] [成功] $Message" -ForegroundColor Green }
        "WARN"    { Write-Host "[$time] [警告] $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "[$time] [错误] $Message" -ForegroundColor Red }
        Default   { Write-Host "[$time] $Message" }
    }
}

function Resolve-RequiredPath([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "指定路径不存在: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

# 确保 7-Zip 及 7zSD.sfx 模块就绪
function Ensure-7ZipTools {
    $toolRoot = Join-Path $env:LOCALAPPDATA 'Excel_Compare\BuildTools\7zip-2601'
    $sevenZip = Join-Path $toolRoot 'app\7z.exe'
    $sfx = Join-Path $toolRoot 'sdk\bin\7zSD.sfx'

    if ((Test-Path -LiteralPath $sevenZip) -and (Test-Path -LiteralPath $sfx)) {
        return [PSCustomObject]@{ SevenZip = $sevenZip; Sfx = $sfx }
    }

    Write-Log "未检测到 7-Zip 打包工具，正在自动初始化工具链..." "INFO"
    New-Item -ItemType Directory -Force -Path $toolRoot | Out-Null
    
    $installer = Join-Path $toolRoot '7z2601-x64.exe'
    $sdk = Join-Path $toolRoot 'lzma2601.7z'

    if (-not (Test-Path -LiteralPath $installer)) {
        Write-Log "正在下载 7-Zip x64 安装程序..." "INFO"
        Invoke-WebRequest -Uri 'https://7-zip.org/a/7z2601-x64.exe' -OutFile $installer
    }

    $appDir = Join-Path $toolRoot 'app'
    New-Item -ItemType Directory -Force -Path $appDir | Out-Null
    $proc = Start-Process -FilePath $installer -ArgumentList @('/S', ('/D=' + $appDir)) -Wait -PassThru -WindowStyle Hidden
    if ($proc.ExitCode -ne 0) { throw "7-Zip 静默安装失败，退出码: $($proc.ExitCode)" }

    if (-not (Test-Path -LiteralPath $sdk)) {
        Write-Log "正在下载 LZMA SDK (包含 SFX 模块)..." "INFO"
        Invoke-WebRequest -Uri 'https://7-zip.org/a/lzma2601.7z' -OutFile $sdk
    }

    $sdkDir = Join-Path $toolRoot 'sdk'
    New-Item -ItemType Directory -Force -Path $sdkDir | Out-Null
    & $sevenZip x $sdk "-o$sdkDir" -y | Out-Null

    if (-not (Test-Path -LiteralPath $sfx)) {
        throw "未能在 SDK 中找到 7zSD.sfx 模块: $sfx"
    }

    Write-Log "7-Zip 工具链初始化完毕！" "SUCCESS"
    return [PSCustomObject]@{ SevenZip = $sevenZip; Sfx = $sfx }
}

# Win32 PE 资源 API 定义（用于提取与注入图标）
function Ensure-ResourceApi {
    if ('SingleFileResApi' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class SingleFileResApi {
    public delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);

    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool FreeLibrary(IntPtr hModule);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool EnumResourceNames(IntPtr hModule, IntPtr lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LockResource(IntPtr hResData);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr BeginUpdateResource(string pFileName, bool bDeleteExistingResources);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool UpdateResource(IntPtr hUpdate, IntPtr lpType, IntPtr lpName, ushort wLanguage, byte[] lpData, uint cbData);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool EndUpdateResource(IntPtr hUpdate, bool fDiscard);
}
'@
}

function Get-ResourceRawBytes([IntPtr]$Module, [IntPtr]$Type, [IntPtr]$Name) {
    $res = [SingleFileResApi]::FindResource($Module, $Name, $Type)
    if ($res -eq [IntPtr]::Zero) { throw "FindResource 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    $size = [SingleFileResApi]::SizeofResource($Module, $res)
    $loaded = [SingleFileResApi]::LoadResource($Module, $res)
    if ($loaded -eq [IntPtr]::Zero) { throw "LoadResource 失败" }
    $ptr = [SingleFileResApi]::LockResource($loaded)
    if ($ptr -eq [IntPtr]::Zero) { throw "LockResource 失败" }
    $bytes = New-Object byte[] $size
    [Runtime.InteropServices.Marshal]::Copy($ptr, $bytes, 0, [int]$size)
    return $bytes
}

# 从源可执行文件中提取内嵌的高清 ICO 图标
function Extract-ExeIcon([string]$ExePath, [string]$OutIconPath) {
    Ensure-ResourceApi
    $module = [SingleFileResApi]::LoadLibraryEx($ExePath, [IntPtr]::Zero, 0x00000002) # LOAD_LIBRARY_AS_DATAFILE
    if ($module -eq [IntPtr]::Zero) { throw "加载可执行文件资源失败: $ExePath" }

    try {
        $groupIds = New-Object System.Collections.Generic.List[int]
        $cb = [SingleFileResApi+EnumResNameProc]{ param($mod, $type, $name, $param)
            $v = $name.ToInt64()
            if (($v -shr 16) -eq 0) { $groupIds.Add([int]$v) }
            return $true
        }
        [void][SingleFileResApi]::EnumResourceNames($module, [IntPtr]14, $cb, [IntPtr]::Zero) # RT_GROUP_ICON = 14
        if ($groupIds.Count -lt 1) {
            Write-Log "源程序未包含 GROUP_ICON 图标资源，将跳过图标提取。" "WARN"
            return $false
        }

        $group = Get-ResourceRawBytes $module ([IntPtr]14) ([IntPtr]$groupIds[0])
        $count = [BitConverter]::ToUInt16($group, 4)
        $images = @()
        $icoLen = 6 + 16 * $count

        for ($i = 0; $i -lt $count; $i++) {
            $entryOff = 6 + 14 * $i
            $iconId = [BitConverter]::ToUInt16($group, $entryOff + 12)
            $img = Get-ResourceRawBytes $module ([IntPtr]3) ([IntPtr]$iconId) # RT_ICON = 3
            $images += ,$img
            $icoLen += $img.Length
        }

        $ico = New-Object byte[] $icoLen
        [Array]::Copy($group, 0, $ico, 0, 6)
        $offset = 6 + 16 * $count

        for ($i = 0; $i -lt $count; $i++) {
            $entryOff = 6 + 14 * $i
            $icoOff = 6 + 16 * $i
            [Array]::Copy($group, $entryOff, $ico, $icoOff, 12)
            [Array]::Copy([BitConverter]::GetBytes([uint32]$offset), 0, $ico, $icoOff + 12, 4)
            [Array]::Copy($images[$i], 0, $ico, $offset, $images[$i].Length)
            $offset += $images[$i].Length
        }

        [System.IO.File]::WriteAllBytes($OutIconPath, $ico)
        return $true
    } finally {
        [void][SingleFileResApi]::FreeLibrary($module)
    }
}

# 将 ICO 图标注入到 SFX 可执行文件中
function Inject-ExeIcon([string]$TargetExePath, [string]$IcoPath) {
    Ensure-ResourceApi
    $icoBytes = [System.IO.File]::ReadAllBytes($IcoPath)
    $count = [BitConverter]::ToUInt16($icoBytes, 4)
    $h = [SingleFileResApi]::BeginUpdateResource($TargetExePath, $false)
    if ($h -eq [IntPtr]::Zero) { throw "BeginUpdateResource 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    $ok = $false

    try {
        $group = New-Object byte[] (6 + 14 * $count)
        [Array]::Copy($icoBytes, 0, $group, 0, 6)
        for ($i = 0; $i -lt $count; $i++) {
            $srcOff = 6 + 16 * $i
            $dstOff = 6 + 14 * $i
            [Array]::Copy($icoBytes, $srcOff, $group, $dstOff, 12)
            $bytesInRes = [BitConverter]::ToUInt32($icoBytes, $srcOff + 8)
            $imageOffset = [BitConverter]::ToUInt32($icoBytes, $srcOff + 12)
            $iconId = [UInt16]($i + 1)
            [Array]::Copy([BitConverter]::GetBytes($iconId), 0, $group, $dstOff + 12, 2)
            $image = New-Object byte[] $bytesInRes
            [Array]::Copy($icoBytes, [int]$imageOffset, $image, 0, [int]$bytesInRes)
            if (-not [SingleFileResApi]::UpdateResource($h, [IntPtr]3, [IntPtr]$iconId, 0, $image, [uint32]$image.Length)) {
                throw "Update RT_ICON 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
            }
        }
        if (-not [SingleFileResApi]::UpdateResource($h, [IntPtr]14, [IntPtr]1, 0, $group, [uint32]$group.Length)) {
            throw "Update RT_GROUP_ICON 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
        }
        $ok = $true
    } finally {
        if (-not [SingleFileResApi]::EndUpdateResource($h, -not $ok)) {
            throw "EndUpdateResource 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
        }
    }
}

# 保守安全瘦身：移除冗余多国语言包与 pdb 文件
function Invoke-SafeFolderClean([string]$Root, [string[]]$Keep) {
    $parent = Split-Path -Parent $Root
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupRoot = Join-Path $parent ("$(Split-Path -Leaf $Root)_瘦身备份_$stamp")
    New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

    $keepSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $Keep | ForEach-Object { [void]$keepSet.Add($_) }
    $knownNonLanguageDirs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    'Assets','Features','FluentIcons.WinUI','Localization','Microsoft.UI.Xaml','winget-cli','runtimes' | ForEach-Object { [void]$knownNonLanguageDirs.Add($_) }

    # 扫描仅包含 .mui / *.resources.dll 的非保留语言目录
    $dirCandidates = Get-ChildItem -LiteralPath $Root -Directory -Force | Where-Object {
        -not $keepSet.Contains($_.Name) -and -not $knownNonLanguageDirs.Contains($_.Name)
    } | Where-Object {
        $files = Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force -ErrorAction SilentlyContinue
        $files.Count -gt 0 -and (($files | Where-Object { $_.Name -notlike '*.mui' -and $_.Name -notlike '*.resources.dll' }).Count -eq 0)
    }

    # 扫描非中英文语言 JSON
    $keepLocalization = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    'en.json','zh-Hans.json','zh-Hant.json','zh-CN.json','zh-TW.json' | ForEach-Object { [void]$keepLocalization.Add($_) }
    $jsonCandidates = @()
    $locDir = Join-Path $Root 'Localization'
    if (Test-Path -LiteralPath $locDir) {
        $jsonCandidates = Get-ChildItem -LiteralPath $locDir -File -Filter '*.json' -Force | Where-Object {
            -not $keepLocalization.Contains($_.Name)
        }
    }

    # 扫描 PDB 调试符号文件
    $pdbCandidates = Get-ChildItem -LiteralPath $Root -Recurse -File -Filter '*.pdb' -Force

    $items = @($dirCandidates) + @($jsonCandidates) + @($pdbCandidates)
    if ($items.Count -eq 0) {
        Write-Log "未发现需要瘦身的冗余文件。" "INFO"
        Remove-Item -LiteralPath $backupRoot -Force -Recurse -ErrorAction SilentlyContinue
        return $null
    }

    $bytes = (($items | ForEach-Object {
        if ($_.PSIsContainer) { (Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force | Measure-Object Length -Sum).Sum } else { $_.Length }
    }) | Measure-Object -Sum).Sum

    foreach ($item in $dirCandidates) {
        $dest = Join-Path $backupRoot ('dirs\' + $item.Name)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Move-Item -LiteralPath $item.FullName -Destination $dest
    }
    foreach ($item in $jsonCandidates) {
        $dest = Join-Path $backupRoot ('Localization\' + $item.Name)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Move-Item -LiteralPath $item.FullName -Destination $dest
    }
    foreach ($item in $pdbCandidates) {
        $rel = $item.FullName.Substring($Root.Length).TrimStart('\')
        $dest = Join-Path $backupRoot ('pdb\' + $rel)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $dest) | Out-Null
        Move-Item -LiteralPath $item.FullName -Destination $dest
    }

    $freedMB = [math]::Round($bytes / 1MB, 2)
    Write-Log "瘦身清理完成：移除 $($dirCandidates.Count) 个语言目录, $($jsonCandidates.Count) 个语言包, $($pdbCandidates.Count) 个PDB文件，释放 $freedMB MB 空间（已备份至 $backupRoot）。" "SUCCESS"

    return [PSCustomObject]@{
        BackupRoot = $backupRoot
        RemovedDirs = $dirCandidates.Count
        RemovedLocalizationJson = $jsonCandidates.Count
        RemovedPdb = $pdbCandidates.Count
        FreedMB = $freedMB
    }
}

# --- 主执行逻辑 ---

$source = Resolve-RequiredPath $SourceDir
$entryPath = if ([System.IO.Path]::IsPathRooted($EntryExe)) { $EntryExe } else { Join-Path $source $EntryExe }
$entryPath = Resolve-RequiredPath $entryPath

if (-not $entryPath.StartsWith($source, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "主入口程序必须位于源文件夹内部: $entryPath"
}

$leaf = Split-Path -Leaf $source
if ([string]::IsNullOrWhiteSpace($OutputName)) {
    $entryLeaf = [System.IO.Path]::GetFileNameWithoutExtension($entryPath)
    $OutputName = "$($entryLeaf)_单文件版"
}

$destDir = if ([string]::IsNullOrWhiteSpace($OutputDir)) { Split-Path -Parent $source } else { $OutputDir }
if (-not (Test-Path -LiteralPath $destDir)) {
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
}
$destDir = (Resolve-Path -LiteralPath $destDir).Path

$finalExe = Join-Path $destDir "$($OutputName).exe"
$tempArchive = Join-Path ([System.IO.Path]::GetTempPath()) "$($OutputName)_$([guid]::NewGuid().ToString('N')).7z"
$tempConfig = Join-Path ([System.IO.Path]::GetTempPath()) "$($OutputName)_$([guid]::NewGuid().ToString('N')).txt"
$extractedIcon = Join-Path $destDir "$($OutputName)_icon.ico"
$tempSfx = Join-Path ([System.IO.Path]::GetTempPath()) "$($OutputName)_$([guid]::NewGuid().ToString('N')).sfx"

Write-Log "==========================================" "INFO"
Write-Log "开始制作单文件版软件: $OutputName" "INFO"
Write-Log "源文件夹: $source" "INFO"
Write-Log "主入口程序: $entryPath" "INFO"
Write-Log "目标输出文件: $finalExe" "INFO"
Write-Log "==========================================" "INFO"

# 1. 瘦身清理（可选）
$cleanInfo = $null
if ($Clean) {
    Write-Log "正在执行打包前目录安全瘦身..." "INFO"
    $cleanInfo = Invoke-SafeFolderClean -Root $source -Keep $KeepCultures
}

# 2. 准备工具
$tools = Ensure-7ZipTools

# 3. 提取或确定图标
$hasIcon = $false
if (-not [string]::IsNullOrWhiteSpace($IconPath) -and (Test-Path -LiteralPath $IconPath)) {
    $iconToUse = Resolve-RequiredPath $IconPath
    $hasIcon = $true
    Write-Log "使用自定义图标: $iconToUse" "INFO"
} else {
    Write-Log "正在提取源程序内嵌的高清图标..." "INFO"
    try {
        $hasIcon = Extract-ExeIcon -ExePath $entryPath -OutIconPath $extractedIcon
        if ($hasIcon) {
            $iconToUse = $extractedIcon
            Write-Log "成功提取图标: $extractedIcon" "SUCCESS"
        }
    } catch {
        Write-Log "图标提取异常: $($_.Exception.Message)" "WARN"
        $hasIcon = $false
    }
}

# 4. 压缩打包文件内容为 7z
Write-Log "正在使用 7-Zip 进行多线程极限压缩 (级别: $CompressionLevel)..." "INFO"
# 将源目录内部的所有文件直接打包到压缩包根目录
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $archiveArgs = @('a', '-t7z', "-mx=$CompressionLevel", '-mmt=on', '-r', $tempArchive, (Join-Path $source '*'))
    Write-Log "执行 7-Zip 命令..." "INFO"
    $proc7z = Start-Process -FilePath $tools.SevenZip -ArgumentList $archiveArgs -NoNewWindow -Wait -PassThru
    $sw.Stop()
    if ($proc7z.ExitCode -ne 0) {
        throw "7-Zip 压缩失败，退出码: $($proc7z.ExitCode)"
    }
Write-Log "压缩完成，耗时: $([math]::Round($sw.Elapsed.TotalSeconds, 2)) 秒，压缩包大小: $([math]::Round((Get-Item $tempArchive).Length / 1MB, 2)) MB" "SUCCESS"

# 5. 生成 SFX 配置文件
$entryRelative = $entryPath.Substring($source.Length).TrimStart('\')
$guiModeVal = switch ($SilentMode) {
    "TempSilent"       { "2" }
    "TempShowProgress" { "1" }
    "CurrentDir"       { "0" }
    Default            { "2" }
}

$configContent = @"
;!@Install@!UTF-8!
Title="$OutputName"
GUIMode="$guiModeVal"
RunProgram="$entryRelative"
;!@InstallEnd@!
"@
[System.IO.File]::WriteAllText($tempConfig, $configContent, [System.Text.UTF8Encoding]::new($false))

# 6. 处理 SFX 模块并注入图标
Copy-Item -LiteralPath $tools.Sfx -Destination $tempSfx -Force
if ($hasIcon -and (Test-Path -LiteralPath $iconToUse)) {
    Write-Log "正在将图标注入到 SFX 单文件可执行体中..." "INFO"
    try {
        Inject-ExeIcon -TargetExePath $tempSfx -IcoPath $iconToUse
        Write-Log "图标注入成功！" "SUCCESS"
    } catch {
        Write-Log "图标注入失败: $($_.Exception.Message)" "WARN"
    }
}

# 7. 拼接合并生成最终单文件 EXE: [SFX Header] + [Config Text] + [7z Archive]
Write-Log "正在合成单文件版可执行文件: $finalExe" "INFO"
Remove-Item -LiteralPath $finalExe -Force -ErrorAction SilentlyContinue

$outStream = [System.IO.File]::Create($finalExe)
try {
    foreach ($part in @($tempSfx, $tempConfig, $tempArchive)) {
        $inStream = [System.IO.File]::OpenRead($part)
        try {
            $inStream.CopyTo($outStream)
        } finally {
            $inStream.Dispose()
        }
    }
} finally {
    $outStream.Dispose()
}

# 8. 校验单文件完整性与自解压结构
Write-Log "正在校验单文件结构完整性..." "INFO"
$verifyOutput = & $tools.SevenZip t $finalExe 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "单文件验证失败！输出: $verifyOutput"
}
Write-Log "单文件完整性校验通过！" "SUCCESS"

# 9. 清理临时产物
Remove-Item -LiteralPath $tempSfx, $tempConfig -Force -ErrorAction SilentlyContinue
if (-not $KeepArchive) {
    Remove-Item -LiteralPath $tempArchive -Force -ErrorAction SilentlyContinue
} else {
    $keepArchiveDest = Join-Path $destDir "$($OutputName).7z"
    Move-Item -LiteralPath $tempArchive -Destination $keepArchiveDest -Force
    Write-Log "已保留 7z 压缩包: $keepArchiveDest" "INFO"
}

$finalSizeMB = [math]::Round((Get-Item $finalExe).Length / 1MB, 2)
Write-Log "==========================================" "SUCCESS"
Write-Log "单文件制作成功！" "SUCCESS"
Write-Log "输出文件: $finalExe ($finalSizeMB MB)" "SUCCESS"
if ($cleanInfo) {
    Write-Log "瘦身备份: $($cleanInfo.BackupRoot) (节省 $($cleanInfo.FreedMB) MB)" "INFO"
}
Write-Log "==========================================" "SUCCESS"

return [PSCustomObject]@{
    SingleFileExe = $finalExe
    SizeMB        = $finalSizeMB
    SourceDir     = $source
    EntryExe      = $entryPath
    IconPath      = if ($hasIcon) { $iconToUse } else { $null }
    CleanBackup   = if ($cleanInfo) { $cleanInfo.BackupRoot } else { $null }
}
