# Excel 自动整理工具

这是一个适用于 Windows 11 的中文桌面工具，使用 Python、tkinter 和 openpyxl。它可以将多个 `.xlsx` 文件中**第一个工作表**的数据合并，按整行去重，按指定列排序，再另存为一个新文件。程序不会修改输入文件。

## 功能

- 一次选择多个 `.xlsx` 文件，按选择顺序合并。
- 检查表头的列名和顺序是否完全一致；遇到空列名、重复列名或不一致的表头会提示错误。
- 删除完全重复的数据行，保留第一次出现的记录。
- 按指定列升序或降序排序，空值始终排在最后。
- 导出新的 `.xlsx` 文件，并拒绝将任一输入文件设为输出路径。

## 界面操作示意

![Excel 自动整理工具界面操作示意图](docs/interface-guide.svg)

此图是按实际界面布局绘制的操作示意，**不是软件截图**。从上到下依次是文件列表、排序列及顺序设置、导出按钮。点击“添加文件”后，排序列下拉框会读取第一个文件的表头。

## 下载与使用 EXE（无需安装 Python）

前往 [v1.0.0 下载页面](https://github.com/wangyichen921/excel-organizer/releases/tag/v1.0.0)，点击显示为“Excel自动整理工具.exe”的附件。GitHub 不保留附件名中的中文字符，实际下载文件名是 `Excel-Organizer-v1.0.0-Windows.exe`；文件内容与本地构建的 `Excel自动整理工具.exe` 相同。双击启动后：

1. 点击“添加文件”，选择两个或更多表头相同的 `.xlsx` 文件。
2. 在“排序列”中选择一列，再选择升序或降序。
3. 点击“整理并另存为…”，选择一个**不同于输入文件**的输出路径。
4. 打开导出的文件，检查结果。

`examples/订单_一.xlsx` 和 `examples/订单_二.xlsx` 可用于试用。选“金额”升序后，输出应有 3 行数据，金额依次是 85、120、200。EXE 不包含示例文件，下载后可从本仓库的 `examples` 目录单独获取。

## 从源代码运行

需要 Windows 11 和 Python 3.10 或更高版本。在 PowerShell 中进入下载或克隆后的项目目录：

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

运行自动化测试：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## 自行打包 EXE

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe -m PyInstaller --onefile --windowed --clean --noconfirm --name Excel自动整理工具 --distpath dist --workpath build packaged_app.py
.\.venv\Scripts\python.exe verify_exe.py
```

生成文件位于 `dist/Excel自动整理工具.exe`，不会提交到源码仓库。`verify_exe.py` 会运行 EXE 的诊断模式，检查合并、排序、输入文件哈希不变，并确认正常模式下的主窗口可见。验证脚本需要 Python；EXE 的普通使用不需要 Python。

## 项目结构

```text
excel-organizer/
├─ app.py                  中文图形界面
├─ organizer.py            表头检查、合并、去重、排序与导出
├─ packaged_app.py         PyInstaller 入口与诊断模式
├─ verify_exe.py           EXE 自动验证脚本
├─ generate_examples.py    示例文件生成脚本
├─ examples/               两份示例输入文件
├─ docs/                   界面操作示意图
├─ tests/test_organizer.py 自动化测试
├─ requirements.txt        运行依赖
└─ requirements-build.txt  打包依赖
```

## 处理范围

程序只读取每个文件的第一个工作表。输出保留单元格值和公式文本，不复制源文件的格式、图片、图表或其他工作表。公式中的跨文件引用和相对引用不会自动重写。去重后的数据保存在内存中处理，大文件需要足够内存。公开仓库目前**未提供开源许可证**。
