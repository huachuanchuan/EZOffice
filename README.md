# EZOffice

EZOffice是一个 Windows 桌面版 Office 加载项，为 **Word、Excel 和 PowerPoint** 提供 **Shift + 滚轮横向滚动**和**本地定时保存**。每个应用都有自己的 **EZOffice** 功能区选项卡，可独立设置两项功能，无需打开侧边栏。

[下载最新版安装包](https://github.com/huachuanchuan/EZOffice/releases/latest/download/EZOffice.zip)

## 安装

1. 关闭 Word、Excel 和 PowerPoint 的所有窗口。
2. 从 [GitHub Releases](https://github.com/huachuanchuan/EZOffice/releases) 下载 `EZOffice.zip`，解压后双击 `Install.cmd`。安装仅适用于当前 Windows 用户。
3. 在 Word、Excel 或 PowerPoint 中打开文件，选择 **EZOffice** 选项卡。

如果应用没有显示选项卡，在 **文件 → 选项 → 加载项** 底部选择 **COM 加载项 → 转到**，勾选 **EZOffice**。卸载时关闭三个应用，再双击安装包里的 `Uninstall.cmd`。

## 使用

- **横向滚动**默认开启。将鼠标指向 Word 页面、Excel 工作表或 PowerPoint 的幻灯片编辑区，按住 **Shift** 滚动滚轮。向上滚往左，向下滚往右。Word 和 PowerPoint 可放大内容以产生横向滚动空间。
- **自动保存**默认关闭。打开开关后，插件定时把该应用内已命名、有修改且可写的本地文件保存回原文件。新文件先按 **Ctrl+S** 选定本地文件名。插件会跳过只读文件、网络位置和检测到的 OneDrive 同步目录。
- **间隔（分钟）**默认为 **5**，可输入 **1–120** 的整数，按 **Enter** 应用。每个应用分别记住开关和间隔，保存结果显示在功能区。

## 兼容性

安装包支持 32 位和 64 位 Windows 桌面版 Office 的 COM 加载项注册。已在 64 位 Office 16.0 中验证三个应用的横向滚动开关、左右滚动，以及按 1 分钟间隔将修改写回 `.docx`、`.xlsx` 和 `.pptx` 原文件。

## 从源码构建

在 Windows PowerShell 中运行 `powershell -NoProfile -ExecutionPolicy Bypass -File .\Build.ps1`。脚本使用系统自带的 .NET Framework 编译器生成 `dist\EZOffice.zip`。

实现参考 [Microsoft 的 COM 功能区加载项示例](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in)，以及 [Word](https://learn.microsoft.com/en-us/office/vba/api/word.window.smallscroll)、[Excel](https://learn.microsoft.com/en-us/office/vba/api/excel.window.smallscroll) 和 [PowerPoint](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.smallscroll) 的滚动接口。

## 许可

本项目采用 [MIT 许可](LICENSE)。
