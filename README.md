# 磁力链接批量处理器 (Magnet Processor)

一个基于 GitHub Actions 的自动化工具，用于批量处理 Excel 文件中的磁力链接，自动获取资源名称、文件数量、总大小以及视频截图，并将结果保存为带截图的 Excel 文件。

## 功能特点

- 🔗 **批量处理**：读取 `Input` 目录下的所有 Excel 文件（`.xlsx`/`.xlsm`/`.xls`），逐行处理磁力链接。
- 📊 **数据提取**：调用 [whatslink.info](https://whatslink.info) API 获取资源名称、文件数、总大小和截图链接。
- 🖼️ **截图嵌入**：自动下载截图并转换为 JPEG，嵌入到输出 Excel 的列中，保留原始清晰度。
- ⏳ **断点续传**：通过进度文件（`.progress.json`）记录已处理行数，中断后可从上次位置继续。
- 📦 **分批输出**：每处理指定行数（默认 20 行）保存一个批次 Excel 文件，避免一次性输出过大。
- 🤖 **自动提交**：每次保存批次后自动提交到 Git（带 `[skip ci]` 标记），便于实时同步。
- ⏱️ **定时运行**：通过 GitHub Actions 每 6 分钟触发一次，自动处理新输入文件，直至全部完成。
- 🔁 **并发避免**：通过检查上次运行时间，避免多个工作流并发冲突。
- 🧹 **临时文件清理**：处理完成后自动删除临时图片文件，保持工作区干净。


## 文件结构

```
├── .github/workflows/magnet-processor.yml   # GitHub Actions 工作流定义
├── Input/                                   # 放置待处理的 Excel 文件（需要手动上传）
│   └── example.xlsx
├── Output/                                  # 生成的批次 Excel 文件和进度文件
│   ├── .progress.json                       # 进度状态（自动生成）
│   ├── 001_example.xlsx
│   ├── 002_example.xlsx
│   └── ...
├── magnet_processor.py                      # 主处理脚本
├── requirements.txt                         # Python 依赖
└── README.md                                # 本文件
```

## 配置说明

所有可调参数位于 `magnet_processor.py` 开头的“用户可调配置”区域，您可以直接修改：

```python
REQUEST_INTERVAL = 0.7          # API 请求间隔（秒），避免触发频率限制
BATCH_SIZE = 20                 # 每批处理行数
MAX_RETRIES = 5                 # 下载图片重试次数
TIMEOUT = 20                    # 请求超时（秒）
DEFAULT_COL_WIDTH = 80          # 截图列宽（字符数）
DEFAULT_ROW_HEIGHT = 274        # 截图行高（磅）
JPEG_QUALITY = 85               # JPEG 压缩质量（1-100）
MAX_RUN_TIME = 300              # 单次最大运行时间（分钟）
MAX_ROWS_PER_RUN = 999999       # 单次运行最大处理行数
```

**工作流级别限制**（在 `.github/workflows/magnet-processor.yml` 中）：
- `timeout-minutes: 310`：工作流最大执行时间 5 小时 10 分钟。
- `cron: '*/6 * * * *'`：每 6 分钟触发一次。

## 使用方法

### 1. 将仓库克隆或 Fork 到您的 GitHub 账号

```bash
git clone https://github.com/your-username/your-repo.git
cd your-repo
```

### 2. 添加输入文件

将包含磁力链接的 Excel 文件放入 `Input` 目录。  
**输入文件格式要求**：
- 第一行为**标题行**（会被忽略）。
- 第一列（A列）必须包含有效的磁力链接（以 `magnet:?xt=urn:btih:` 开头）。
- 其他列（B、C、D等）内容可任意，但程序将覆盖 B（资源名）、C（文件数）、D（总大小），并在 E 列开始插入截图。
- 支持 `.xlsx`、`.xlsm`、`.xls` 格式。

示例：

| A (磁力链接) | B (原名称) | C (原文件数) | D (原大小) |
|--------------|------------|--------------|------------|
| magnet:?xt=... | 任意 | 任意 | 任意 |
| magnet:?xt=... | ... | ... | ... |

### 3. 提交并推送

```bash
git add Input/
git commit -m "添加待处理的磁力链接文件"
git push
```

### 4. 自动触发

推送后，GitHub Actions 会自动运行（若工作流配置为 `on: push` 或定时触发）。您可以手动触发工作流：

- 进入仓库的 **Actions** 选项卡
- 选择 **磁力链接批量处理** 工作流
- 点击 **Run workflow** → **Run workflow**

### 5. 查看输出

处理完成后，生成的批次文件（`Output/*.xlsx`）会自动提交到仓库，并作为 **Artifacts** 保留 30 天。您可以在 Actions 运行记录的 **Artifacts** 区域下载所有输出文件。

## 输出说明

- **批次文件命名**：`{序号}_{原始文件名}.xlsx`，例如 `001_example.xlsx`。
- **内容结构**：
  - 第 1 行：标题行（来自输入文件的第一行）。
  - 从第 2 行开始：数据行，包含：
    - A列：磁力链接
    - B列：资源名称
    - C列：文件数量
    - D列：总大小（自动转换为 KB/MB/GB/TB）
    - E列及以后：每个截图（嵌入图片，列宽自动调整）
- **进度文件**：`.progress.json` 记录每个文件的处理进度（已处理行数、完成状态等），用于断点续传。

## 注意事项

- **API 限制**：whatslink.info API 可能有频率限制，请合理设置 `REQUEST_INTERVAL`（默认 0.7 秒）。
- **网络依赖**：处理过程需要访问外网下载截图，请确保 GitHub Actions 运行环境网络通畅。
- **Git 提交**：每次保存批次后会执行 `git add` 和 `git commit`，并自动推送。如需关闭，可注释掉 `git_commit_files` 调用。
- **运行时长**：单个工作流最长运行 5 小时，若超过会强制终止。您可以在 `MAX_RUN_TIME` 和 `timeout-minutes` 中调整。
- **输入文件变动**：若 `Input` 中的文件被修改（哈希变化），进度会重置，重新处理。
- **并发处理**：工作流通过检查 `.progress.json` 的更新时间，避免多个实例同时运行（间隔小于 5 分钟则跳过）。

## 依赖项

项目使用以下 Python 库（在 `requirements.txt` 中指定）：

- `requests` – API 调用
- `openpyxl` – 读取 Excel
- `xlsxwriter` – 写入 Excel（支持图片嵌入）
- `Pillow` – 图片格式转换

这些依赖会在 GitHub Actions 运行时自动安装。

## 许可证

[MIT](LICENSE)
