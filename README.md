# 文档处理系统

一个基于Python的智能文档处理系统，支持Word文档(.docx/.doc)和PDF文档的自动处理，将其转换为Markdown格式并提取关键元数据。

## 功能特性

- **多格式支持**: 支持Word文档(.docx/.doc)和PDF文档处理
- **智能转换**: 自动将文档转换为结构化的Markdown格式
- **元数据提取**: 自动生成文档摘要和关键词
- **OCR处理**: 使用PaddleOCR对PDF文档进行高精度文字识别
- **批量处理**: 支持目录下多个文档的批量处理

## 项目结构

```
doc_preparation/
├── config.py              # 配置文件
├── main.py                # 主程序入口
├── requirements.txt       # 依赖包列表
├── README.md              # 项目说明
├── input/                 # 输入文档目录
├── output/                # 处理结果输出目录
├── models/                # 预训练模型缓存目录
├── custom_models/         # 自定义OCR模型目录
└── core/                  # 核心功能模块
    ├── __init__.py
    ├── metadata_extractor.py  # 元数据提取模块
    ├── utils.py               # 工具函数
    └── converters/            # 文档转换器
        ├── __init__.py
        ├── docx_converter.py  # Word文档转换器
        └── pdf_converter.py   # PDF文档转换器
```

## 安装与配置

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置参数

在 `config.py` 中根据需要调整配置参数：

```python
# 基础目录配置
INPUT_DIR = BASE_DIR / "input"    # 输入文档目录
OUTPUT_DIR = BASE_DIR / "output"  # 输出目录

# 文档处理参数
SUMMARY_SENTENCES = 3             # 摘要句子数量
KEYWORDS_TOP_N = 10              # 关键词提取数量

# 模型配置
USE_LOCAL_MODELS = True          # 是否使用本地模型
USE_GPU_FOR_OCR = True          # 是否使用GPU进行OCR
```

## 使用方法

### 1. 准备文档

将需要处理的Word文档(.docx/.doc)或PDF文档放入 `input/` 目录。

### 2. 运行处理程序

```bash
python main.py
```

### 3. 查看结果

处理完成后，结果将保存在 `output/` 目录中：

- `文档名.md`: Markdown格式的文档内容
- `文档名_metadata.json`: 提取的元数据（摘要、关键词等）

## 输出示例

### Markdown文件
处理后的文档将转换为结构化的Markdown格式，保留原文档的层次结构、表格和图片。

### 元数据文件
```json
{
  "summary": "文档摘要内容...",
  "keywords": [
    ["关键词1", 0.85],
    ["关键词2", 0.72],
    ...
  ],
  "char_count": 15632,
  "word_count": 2341,
  "title": "文档标题",
  "author": "作者",
  "created": "2025-01-01T00:00:00",
  "modified": "2025-01-02T00:00:00"
}
```

## 依赖说明

- **python-docx**: Word文档处理
- **paddleocr**: PDF文档OCR处理
- **tqdm**: 进度条显示
- **sumy**: 文本摘要生成
- **keybert**: 基于BERT的关键词提取
- **jieba**: 中文分词
- **transformers**: Transformer模型库

## 注意事项

1. 首次运行时会自动下载所需的预训练模型
2. PDF处理需要较大的计算资源，建议使用GPU加速
3. 确保输入文档编码正确，避免乱码问题
4. 大批量处理时建议分批进行，避免内存溢出

## 故障排除

### 模型下载失败
如果遇到模型下载问题，可以：
1. 设置 `USE_LOCAL_MODELS = False` 允许在线下载
2. 手动下载模型到 `models/` 目录
3. 检查网络连接和代理设置

### OCR处理失败
如果PDF处理出现问题：
1. 检查CUDA环境配置（GPU模式）
2. 设置 `USE_GPU_FOR_OCR = False` 使用CPU模式
3. 确保custom_models目录包含所需的OCR模型

### 内存不足
处理大文档时如遇内存问题：
1. 减少批处理文档数量
2. 降低图片分辨率
3. 使用更大内存的机器