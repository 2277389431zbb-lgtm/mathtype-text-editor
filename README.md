# MathType 文本编辑器

Windows 桌面工具：粘贴包含正文和 LaTeX 公式的整段内容，自动把公式转换为可编辑的 MathType 对象；在软件中编辑完成后，将文字和公式一起插入 Microsoft Word。

## 下载使用

下载 [Windows v2.0.0 压缩包](downloads/MathTypeTextEditor-v2.0.0-win-x86.zip)，解压后运行 `MathTypeTextEditor.exe`。

需要安装：

- Windows 与 .NET Framework 4.8。
- Microsoft Word，且 Word 中的 MathType 7 加载项可用。
- MathType 7。本项目不包含 Word、MathType 安装程序、第三方字体或许可证。

## 操作流程

1. 从 Word 或浏览器复制整段文字，在软件编辑区按 **Ctrl+V**。
2. 默认自动转换其中的公式，同时保留剪贴板提供的常见文字格式。
3. 继续编辑文字、字号、加粗、颜色和段落；双击公式可以打开 MathType 编辑。
4. 在目标 Word 文档中放好光标，在软件底部选择该文档，点击 **插入回 Word**。

支持草稿保存、自动恢复、批量转换、撤销最近一批转换，以及选择多个已打开的 Word 文档。

## 公式识别

- 行内公式：`\( ... \)`、`$ ... $`。
- 独立公式：`\[ ... \]`、`$$ ... $$`、`equation` 环境。
- LaTeX 代码块、方括号分别占行的常见复制格式，以及含明确 TeX 命令的普通括号表达式。
- 浏览器富文本中的 KaTeX、MathJax 源码及常见 MathML 结构。
- 已有 MathType 对象保持可编辑，不会再次切换为 TeX。

`\tag{5-24}` 会保留为公式后方的编号文字。转换失败的公式保留原代码，并给出提示。

## 格式与使用范围

保留富文本剪贴板中的字体、字号、加粗、斜体、下划线、颜色和常见段落格式。转换只替换公式片段，不重新生成其他文字。

纯文本剪贴板不包含原字号、颜色等信息，因此不能恢复这些源格式。公式截图不支持直接识别。复杂页面布局、分节、页眉页脚、脚注和复杂表格请继续在 Word 中处理。Word 原生公式不是本版 TeX 扫描的转换对象。

在 Word 存在选区时，程序插在选区末尾，不自动覆盖选区。插入后可在 Word 中按 **Ctrl+Z** 撤销。

详细操作见 [使用说明](docs/使用说明.txt)。

## 从源码构建

使用 Windows 自带的 .NET Framework C# 编译器，不需要 NuGet、Python 或 Node.js。

```powershell
.\build.ps1
```

输出位于 `dist/`。运行解析自检：

```powershell
.\dist\MathTypeTextEditor.exe --self-test
```

安装好 Word 与 MathType 后，可运行完整集成检查：

```powershell
.\dist\MathTypeTextEditor.exe --integration-test
```

检查结果写入可执行文件所在文件夹，集成检查使用独立的后台 Word 实例。

## 项目结构

```text
src/           富文本编辑器、公式识别、Word 自动化和测试源码
docs/          使用说明与验证范围
downloads/     可直接使用的 Windows 压缩包与 SHA-256 校验
build.ps1      本地构建脚本
```

本地草稿存储在 `%LOCALAPPDATA%\MathTypeOneClick\draft.rtf`。仓库不包含个人草稿、毕业设计文档或运行日志。
