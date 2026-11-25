# CATIA V5 R2018 实体拆分 · V02

将一个多实体 CATPart 按几何体名称拆成多个 CATPart，并生成引用这些零件的 CATProduct。不同名的实体分别导出，同名实体放入同一个 CATPart，文件内各自独立。

## 运行环境

- Windows x64 和 .NET 10 Desktop Runtime（x64）
- 已启动的 CATIA V5 R2018（R28）
- 需要 DeepSeek 翻译时，准备可访问的接口和有效的 API Key

## 使用

1. 从 CATIA V5 R2018 的快捷方式启动 CATIA。如果同时开着 R2016，请关闭 R2016 后再运行拆分工具。
2. 双击发布版 `CatPartSplitter_V02.exe`，选择自己的源 CATPart，并选择或输入输出目录。目录不存在时，程序会创建它。若从源码获取项目，请先按下文发布 EXE。
3. 如果有需要 DeepSeek 翻译的名称，在窗口输入 API Key。使用官方接口时保留默认 API 地址 `https://api.deepseek.com`。
4. 点击 **拆分并导出**。完成后，在新生成的子目录中打开 `*_Split.CATProduct`。

## API 配置

首次成功导出后，程序将 Key 明文写入 EXE 同目录的 `api.txt`，将 API 地址写入 `api-url.txt`。下次启动会自动读取。更换 Key 时，在窗口修改并重新导出即可。请勿共享 `api.txt`。

## 英文文件名和术语库

程序优先查找 [空调零部件中英术语库](空调零部件中英术语库.md)，用完整中文名称匹配表格，采用「英文」列的写法。单一译名可在本地直接确定。遇到多条不同译名或没有完整匹配时，DeepSeek 会参考表内候选或相关术语。CATIA 内部名称继续沿用原名称。

公开仓库提供空白模板。填写自己的术语并重新构建发布后，EXE 会内置你填写的术语。

## 输出和拆分规则

| 文件 | 内容 |
| --- | --- |
| `*.CATPart` | 文件名使用英文，CATIA 零件编号使用原几何体名称。同名实体保存在同一零件中。 |
| `*_Split.CATProduct` | 装配实例使用原几何体名称加 `.1`，各组件沿用源文件里的位置。 |
| `split-manifest.json` | 原几何体名称、英文文件名和实体数量。 |

程序处理 CATPart 下直接可枚举的实体几何体，跳过空几何体和已参与布尔运算的操作体。输出采用无链接的几何结果，保留形状和位置。若要移动或共享总装，请将整个输出子目录一起移动。

每个导出零件的 XY、YZ、ZX 原点平面默认隐藏，仍可在 CATIA 树上切换显示。每次导出随机生成配色，同一 CATPart 内的实体共用一种颜色。

## 构建与检查

在此目录使用 .NET 10 SDK 构建。

```powershell
dotnet build .\CatPartSplitter\CatPartSplitter.vbproj -c Release
```

重新发布单文件 EXE 时运行以下命令。

```powershell
dotnet publish .\CatPartSplitter\CatPartSplitter.vbproj -c Release -r win-x64 --no-self-contained -p:PublishSingleFile=true -p:EnableSingleFileAnalyzer=false -p:DebugType=None -o .\CatPartSplitter\bin\publish-v02
Copy-Item -LiteralPath .\CatPartSplitter\bin\publish-v02\CatPartSplitter.exe -Destination .\release\CatPartSplitter_V02.exe
```

也可以在命令行检查输入中的分组，检查无需联网。

```powershell
.\release\CatPartSplitter_V02.exe "D:\CAD\YourPart.CATPart" --inspect
```

命令行导出可用 `--output` 指定输出目录，通过环境变量 `DEEPSEEK_API_KEY` 提供 Key。若已准备 UTF-8 JSON 名称映射，可用 `--name-map names.json` 指定。映射格式为原几何体名称到英文名称，例如 `{"原名称":"English Name"}`。手工映射优先于术语库，两者覆盖所有分组时可离线导出。

## GitHub 源码包

源码包含 VB.NET 项目、图标和构建所需的 Markdown 术语库空白模板。运行时请使用自己的 CATPart，首次需要联网翻译时在界面填写自己的 Key。`api.txt`、`api-url.txt`、CAD 文件、导出清单及构建目录均由 `.gitignore` 排除。

## 常见问题

- **提示找不到 R2018 会话**。确认已启动 R2018，且程序与 CATIA 在同一 Windows 用户会话下运行。若开着 R2016，先关闭 R2016 再试。
- **DeepSeek 翻译失败**。检查 API Key、HTTPS 地址和网络连接。官方地址为 `https://api.deepseek.com`。
- **下次启动没有自动加载 Key**。确认上次成功导出，且 EXE 所在目录允许当前用户创建 `api.txt`。
- **总装打不开零件链接**。确认导出的 CATPart 都在总装同一输出目录中，文件名与导出时一致。
