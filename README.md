# ppt-to-images

将 PPT/PPTX/PDF 文件转换为图片序列的 Python 工具。

## 功能特性

- 支持 PPT、PPTX、PDF 三种输入格式
- 支持文件输出、Base64 输出、或两者同时输出
- 可选提取幻灯片备注内容
- 命令行工具和 Python API 两种使用方式
- 可配置输出 DPI（分辨率）

## 系统依赖

本工具依赖以下系统软件：

- **LibreOffice** - 用于 PPT/PPTX 转 PDF
- **Poppler** - 用于 PDF 转图片（pdf2image 依赖）

### Ubuntu/Debian

```bash
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

### macOS

```bash
brew install libreoffice poppler
```

### Windows

1. 下载并安装 [LibreOffice](https://www.libreoffice.org/download/download/)
2. 下载并安装 [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases)

## 安装

### 从源码安装

```bash
git clone https://github.com/HQIT/ppt-to-images.git
cd ppt-to-images
pip install -e .
```

### 使用 pip 安装

```bash
pip install ppt-to-images
```

## 命令行使用

### 基本用法

```bash
# 将 PPTX 转换为图片，保存到指定目录
ppt-to-images input.pptx -o ./output/images

# 转换 PDF 文件
ppt-to-images input.pdf -o ./output/images

# 转换 PPT 文件
ppt-to-images input.ppt -o ./output/images
```

### 输出格式

```bash
# 文件输出（默认）
ppt-to-images input.pptx -o ./images --format file

# Base64 输出（输出到标准输出）
ppt-to-images input.pptx --format base64

# 同时输出文件和 Base64
ppt-to-images input.pptx -o ./images --format both
```

### 高级选项

```bash
# 设置输出 DPI（默认 200）
ppt-to-images input.pptx -o ./images --dpi 300

# 提取备注内容
ppt-to-images input.pptx -o ./images --extract-notes

# JSON 格式输出
ppt-to-images input.pptx -o ./images --output-json

# 详细输出
ppt-to-images input.pptx -o ./images -v

# 保留临时文件（调试用）
ppt-to-images input.pptx -o ./images --keep-temp

# 指定临时目录
ppt-to-images input.pptx -o ./images --temp-dir /tmp/ppt-convert
```

### 完整示例

```bash
ppt-to-images presentation.pptx \
    -o ./output/images \
    --format both \
    --dpi 300 \
    --extract-notes \
    --output-json \
    -v
```

### 命令行参数

| 参数 | 说明 |
|------|------|
| `input` | 输入文件路径（必需） |
| `-o, --output-dir` | 输出目录 |
| `-f, --format` | 输出格式：file/base64/both（默认 file） |
| `--dpi` | 图片 DPI（默认 200） |
| `--extract-notes` | 提取备注内容 |
| `--output-json` | JSON 格式输出 |
| `--temp-dir` | 临时文件目录 |
| `--keep-temp` | 保留临时文件 |
| `-v, --verbose` | 详细输出 |
| `--version` | 显示版本号 |

## Python API

### 基本用法

```python
from ppt_to_images import PPTConverter

# 创建转换器
converter = PPTConverter(dpi=200)

# 转换并保存为文件
result = converter.convert(
    input_path="input.pptx",
    output_dir="./output/images",
    output_format="file"
)

print(f"转换了 {result['count']} 页")
print(f"图片路径: {result['images']}")
```

### 输出为 Base64

```python
from ppt_to_images import PPTConverter

converter = PPTConverter()

result = converter.convert(
    input_path="input.pptx",
    output_format="base64"
)

# result["images"] 包含 Base64 编码的图片列表
for i, b64 in enumerate(result["images"], 1):
    print(f"Page {i}: {b64[:50]}...")
```

### 同时输出文件和 Base64

```python
from ppt_to_images import PPTConverter

converter = PPTConverter()

result = converter.convert(
    input_path="input.pptx",
    output_dir="./output",
    output_format="both",
    extract_notes=True
)

print(f"文件路径: {result['images']}")
print(f"Base64: {result['images_base64']}")
print(f"备注内容: {result['texts']}")
```

### 处理 PDF 文件

```python
from ppt_to_images import PPTConverter

converter = PPTConverter(dpi=300)

result = converter.convert(
    input_path="document.pdf",
    output_dir="./output",
    output_format="file"
)
```

### API 参考

#### PPTConverter

```python
PPTConverter(
    dpi: int = 200,           # 输出图片 DPI
    cleanup_temp: bool = True, # 是否清理临时文件
    temp_dir: str = None       # 自定义临时目录
)
```

#### convert()

```python
converter.convert(
    input_path: str | Path,    # 输入文件路径
    output_dir: str = None,    # 输出目录
    output_format: str = "file", # 输出格式：file/base64/both
    extract_notes: bool = False   # 是否提取备注
) -> dict
```

返回值：

```python
{
    "images": list,       # 文件路径或 Base64 列表
    "images_base64": list, # Base64 列表（仅 format="both"）
    "count": int,         # 图片数量
    "texts": list,        # 备注内容（仅 extract_notes=True）
    "format": str         # 输出格式
}
```

## 作为模块运行

```bash
python -m ppt_to_images input.pptx -o ./output
```

## License

MIT License

