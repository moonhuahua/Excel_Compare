#Requires AutoHotkey v1.1.33+
#NoEnv
#SingleInstance Force
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

; 可直接用两个参数调用；不带参数运行时，注册资源管理器专用热键。
if (A_Args.Length() = 2)
{
    CompareExcel(A_Args[1], A_Args[2])
    ExitApp
}
if (A_Args.Length() != 0)
{
    MsgBox, 48, Excel 对比, 请传入恰好两个文件路径。
    ExitApp, 1
}
return

; 在资源管理器中选中两个文件，按 Ctrl+Alt+E。
#If WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass")
^!e::
    files := ExplorerSelectedFiles()
    if (files.Length() != 2)
    {
        MsgBox, 48, Excel 对比, 请先在同一个资源管理器窗口选中恰好两个 Excel 文件。
        return
    }
    CompareExcel(files[1], files[2])
return
#If

ExplorerSelectedFiles()
{
    files := []
    activeHwnd := WinExist("A")
    shell := ComObjCreate("Shell.Application")
    for window in shell.Windows
    {
        try
        {
            if (window.HWND != activeHwnd)
                continue
            for item in window.Document.SelectedItems
                files.Push(item.Path)
            break
        }
    }
    return files
}

CompareExcel(leftFile, rightFile)
{
    exePath := A_ScriptDir . "\Excel_Compare.exe"
    if !FileExist(exePath)
    {
        MsgBox, 48, Excel 对比, % "找不到程序：`n" . exePath . "`n`n请将本脚本与 EXE 放在同一目录。"
        return false
    }
    for index, filePath in [leftFile, rightFile]
    {
        attributes := FileExist(filePath)
        if (!attributes || InStr(attributes, "D"))
        {
            MsgBox, 48, Excel 对比, % "文件不存在或所选项目是文件夹：`n" . filePath
            return false
        }
        SplitPath, filePath,,, extension
        if extension not in xls,xlsx,xlsm,csv,tsv
        {
            MsgBox, 48, Excel 对比, % "不支持的文件类型：`n" . filePath
            return false
        }
    }
    quote := Chr(34)
    command := quote . exePath . quote . " " . quote . leftFile . quote . " " . quote . rightFile . quote
    Run, %command%, %A_ScriptDir%, UseErrorLevel, launchedPid
    if ErrorLevel
    {
        MsgBox, 48, Excel 对比, 无法启动 Excel_Compare，请检查程序文件。
        return false
    }
    return true
}
