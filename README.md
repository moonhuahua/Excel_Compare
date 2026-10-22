<img src="assets/Excel_Compare.png" width="80" height="80" alt="Excel_Compare 图标" />

# Excel_Compare

Windows Excel 双栏对比工具，基于 [skanmera/ExcelMerge](https://github.com/skanmera/ExcelMerge) 修改，保留上游 MIT 许可。

## 下载与运行

从 [Releases](https://github.com/moonhuahua/Excel_Compare/releases/latest) 下载 `Excel_Compare.exe` 单文件，或下载含说明与 AHK 脚本的 `Excel_Compare-Windows.zip`。解压便携包后双击 EXE 即可运行。

需要 Windows 和 .NET Framework 4.5.2 或兼容的 4.x 运行时。无需安装 Microsoft Excel。单文件运行时会解包至临时目录，退出后清理。

## 功能

- 左右独立选择文件和工作表，支持 XLSX、XLS、CSV、TSV；XLSM 兼容性范围见使用说明。
- 差异定位、只显示差异、搜索、同步滚动及可折叠差异面板。
- 自动换行、保留常见单元格格式、可选暗黑模式。
- 列宽随窗口和差异面板调整，支持手动拖动列宽。
- 输入文件只读，命令行传入两个文件后自动执行对比。

```text
Excel_Compare.exe "原文件.xlsx" "新文件.xlsx"
```

便携包中的 `Excel选中两文件对比.ahk` 使用 AutoHotkey **v1.1.33+**：运行后在资源管理器选中两个文件，按 `Ctrl+Alt+E` 对比；也可直接向脚本传入两个路径。脚本与 EXE 应放在同一目录。

## 构建和验证

Windows 上安装 Python 3、.NET SDK 8，运行：

```powershell
python restore_dependencies.py
powershell -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1 -Test
```

依赖按固定版本恢复，首次需要联网；单文件打包脚本随源码提供，不依赖个人 Codex 配置。`-Test` 运行 17 项核心比较检查和 57 项显示检查。输出位于 `发布`，依赖缓存与构建产物不纳入 Git。

详见[使用说明](使用说明.md)。比较以原引擎读取的值为准，公式使用缓存结果，不重算公式、不写回工作簿；显示保留格式不代表进行格式差异比较。
