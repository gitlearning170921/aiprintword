# 批量打印工具

在批量打印 Word、Excel、PDF 前自动完成：

- **Word**：检查封面签字与修订。表格仅做保守修复（取消固定行高），避免字段位置错位与页码波动。
- **Excel**：矩阵外清色；表头底色/框线仅处理有内容格；规范化与打印前会**自适应行高**。每个工作表的 **`PrintArea` 会与 `UsedRange` 对齐**，保证**整本工作簿**各表内容都进入打印范围；**不修改纸张方向**（横向/纵向沿用源文件）。打印前另做水平居中、多列时一页宽等（不改方向）。
- **PDF**：直接提交到默认或指定打印机（依赖系统关联的 PDF 程序，如 Edge/Adobe）。

## 环境要求

- **Windows**（依赖系统打印与 COM）
- **Python 3.8+**
- **默认使用 WPS 文字 / WPS 表格**（KWPS、KET）进行 .doc/.docx/.xls/.xlsx 的检查、修订与打印；未安装 WPS 时会报错
- 若需改用 **Microsoft Word / Excel**，请设置环境变量：`USE_OFFICE=1` 后再运行
- PDF 打印依赖系统默认与 `.pdf` 关联的程序

## 安装

```bash
cd aiprintword
pip install -r requirements.txt
```

## 使用

### Web 界面（推荐）

在项目目录下启动服务后，在浏览器中打开页面即可上传文件、配置选项并执行批量打印：

```bash
python app.py
```

**Windows**：也可双击 **`start_server.bat`** 启动（自动进入脚本所在目录；若存在 **`.venv`** 或 **`venv`** 则优先用其中的 `python.exe`，否则使用系统 `python`）。

浏览器访问 **http://127.0.0.1:5050**（本机）或 **http://\<本机IP\>:5050**（局域网）。

- **选择文件**：支持多选 .doc / .docx / .xls / .xlsx / .pdf
- **打印选项**：勾选是否检查封面签字、是否自动接受修订、是否仅预检不打印；选择打印机与份数
- **开始检查并打印**：流式进度条按步骤估算进度（单文件不会一开始就满格）；结束后汇总「已提交打印队列」或「预检通过」等说明，与每行状态一致

上传的文件会在服务端临时保存并处理，打印完成后自动删除，不会保留在服务器。

### 在线签名（Word / Excel，可选）

与批量打印**独立**：在浏览器打开 **http://127.0.0.1:5050/sign**（或首页「在线签名」链接），上传单个 **.docx** 或 **.xlsx**，按角色手写**签名**与**日期**（均为手写板图片），提交后下载带 `_signed` 后缀的文件。服务端使用 `python-docx` / `openpyxl` 将图片写入模板中已预留的空白位，**不经过 WPS/COM**，也不影响首页批量打印接口。

### 命令行

```bash
# 打印当前目录下所有支持的文档（不递归子目录）
python batch_print.py .

# 递归目录下所有 Word/Excel/PDF
python batch_print.py D:\文档文件夹 -r

# 仅检查 + 接受修订，不实际打印（试跑）
python batch_print.py D:\文档文件夹 -r --dry-run

# 不检查封面签字、不自动接受修订（有修订会跳过）
python batch_print.py file1.docx file2.xlsx --no-signature-check --no-accept-revisions

# 指定打印机与份数
python batch_print.py . -p "Microsoft Print to PDF" -n 2
```

### 一键提交并推送（Windows）

根目录 **`commit_push.bat`**：执行 `git add -A` → `commit`（默认说明 `chore: update`）→ `push origin main`。可传参作为提交说明，例如：`commit_push.bat fix: 进度条`。

**说明**：脚本内提示为英文，避免在 UTF-8 编码下 `cmd.exe` 误解析中文导致 `'IT_MSG'` 等乱报错；结束时仍会 `pause` 方便查看日志。

**`repush.bat`**：只做 **`git push -u origin <当前分支>`**（不 `add`、不 `commit`），用于本地已提交后重新推送到远程或重试推送。

### 参数说明

| 参数 | 说明 |
|------|------|
| `paths` | 一个或多个文件或目录路径 |
| `-r` / `--recursive` | 对目录递归查找文档 |
| `--no-signature-check` | 不检查 Word 封面电子签字 |
| `--no-accept-revisions` | 不自动接受修订；有修订时该文档会被跳过 |
| `-p` / `--printer` | 打印机名称，不填则用系统默认打印机 |
| `-n` / `--copies` | 打印份数，默认 1 |
| `--dry-run` | 只做签字检查与接受修订，不真正打印 |

### 在代码中调用

```python
from batch_print import run_batch

result = run_batch(
    [r"D:\文档目录"],
    recursive=True,
    check_signature=True,
    accept_revisions=True,
    printer_name=None,
    copies=1,
    dry_run=False,
)
print(result["ok"], result["failed"], result["details"])
```

## 行为说明

1. **封面签字（仅 Word）**  
   通过封面页（第一页）检查 **作者、审核、批准、日期** 是否均已填写。若任一项为空或仅为占位符（空格、下划线、横线等），则视为签字未完成，该文档不打印并报错。

2. **字体统一为黑色（Word / Excel）**  
   打印前自动将文档中所有文字颜色统一为黑色，符合正式文档输出要求。会直接修改并保存原文件。

3. **修订（Word）**  
   若检测到存在“修订”（跟踪的更改），且未使用 `--no-accept-revisions`，则会在**原文件上**接受所有修订并保存，然后提交打印。  
   若使用 `--no-accept-revisions` 且文档有修订，则跳过打印并提示。

4. **修订（Excel）**  
   若工作簿为共享并启用跟踪修订，会尝试接受所有更改并保存后再打印；非共享工作簿无修订则直接打印。

5. **PDF**  
   不做签字/修订检查，直接调用系统“打印”动词（通常使用与 PDF 关联的默认程序）。

## 注意事项

- Word/Excel 的检查与打印依赖本机已安装的 Office 及 COM，需在 Windows 下运行。
- 自动接受修订会**修改并保存**原 Word/Excel 文件，建议先备份或先用 `--dry-run` 试跑。
- 打印机名称需与系统中显示的完全一致，可在“设置 → 打印机”中查看。
- **Web 模式**：打印在**运行 Web 服务的电脑**上执行（即安装 WPS/Office 的那台机器）。若在另一台电脑的浏览器中打开页面，文件会传到服务器所在电脑进行打印。
- **WPS 与 Office**：默认调用 WPS（避免 Office RPC 不可用）。使用前请安装 [WPS 办公](https://www.wps.cn/)。要改回 Microsoft Office，在启动前设置 `USE_OFFICE=1`（如 `set USE_OFFICE=1` 后运行 `python app.py`）。

## 故障排除

### 报错 `AttributeError: ... has no attribute 'CLSIDToClassMap'`

这是 pywin32 的 **gencache 缓存**与已安装的 Office 类型库冲突导致的（例如本机装过 Word，缓存了 Word 的 CLSID）。本工具已改为使用 **dynamic.Dispatch** 调用 WPS/Office，不再依赖 gencache，一般可避免该错误。

若仍想单独验证 WPS COM 是否可用，请用**动态调度**测试（不要用 `Dispatch`）：

```bash
python -c "import win32com.client; w=win32com.client.dynamic.Dispatch('KWPS.Application'); print('WPS COM OK'); w.Quit()"
```

若上述命令能正常输出 `WPS COM OK`，说明 WPS 已正确注册，可正常使用本工具。
