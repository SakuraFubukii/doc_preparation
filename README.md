# Document Processing System

[English](README.md) | [中文](README_zh.md)

An intelligent Python-based document processing system that supports automatic processing of Word documents (.docx/.doc) and PDF documents, converting them to Markdown format and extracting key metadata.

## Features

- **Multi-format Support**: Support for Word documents (.docx/.doc) and PDF document processing
- **Intelligent Conversion**: Automatically convert documents to structured Markdown format
- **Metadata Extraction**: Automatically generate document summaries and keywords
- **OCR Processing**: High-precision text recognition for PDF documents using PaddleOCR
- **Batch Processing**: Support for batch processing of multiple documents in a directory

## Project Structure

```text
doc_preparation/
├── config.py              # Configuration file
├── main.py                # Main program entry point
├── requirements.txt       # Dependencies list
├── README.md              # Project documentation (English)
├── README_zh.md           # Project documentation (Chinese)
├── input/                 # Input documents directory
├── output/                # Processing results output directory
├── models/                # Pre-trained models cache directory
├── custom_models/         # Custom OCR models directory
└── core/                  # Core functionality modules
    ├── __init__.py
    ├── metadata_extractor.py  # Metadata extraction module
    ├── utils.py               # Utility functions
    └── converters/            # Document converters
        ├── __init__.py
        ├── docx_converter.py  # Word document converter
        └── pdf_converter.py   # PDF document converter
```

## Installation and Configuration

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Parameters

Adjust configuration parameters in `config.py` as needed:

```python
# Basic directory configuration
INPUT_DIR = BASE_DIR / "input"    # Input documents directory
OUTPUT_DIR = BASE_DIR / "output"  # Output directory

# Document processing parameters
SUMMARY_SENTENCES = 3             # Number of summary sentences
KEYWORDS_TOP_N = 10              # Number of keywords to extract

# Model configuration
USE_LOCAL_MODELS = True          # Whether to use local models
USE_GPU_FOR_OCR = True          # Whether to use GPU for OCR
```

## Usage

### 1. Prepare Documents

Place Word documents (.docx/.doc) or PDF documents to be processed in the `input/` directory.

### 2. Run Processing Program

```bash
python main.py
```

### 3. View Results

After processing is complete, results will be saved in the `output/` directory:

- `document_name.md`: Document content in Markdown format
- `document_name_metadata.json`: Extracted metadata (summary, keywords, etc.)

## Output Examples

### Markdown Files

Processed documents will be converted to structured Markdown format, preserving the original document's hierarchy, tables, and images.

### Metadata Files

```json
{
  "summary": "Document summary content...",
  "keywords": [
    ["keyword1", 0.85],
    ["keyword2", 0.72],
    ...
  ],
  "char_count": 15632,
  "word_count": 2341,
  "title": "Document Title",
  "author": "Author",
  "created": "2025-01-01T00:00:00",
  "modified": "2025-01-02T00:00:00"
}
```

## Dependencies

- **python-docx**: Word document processing
- **paddleocr**: PDF document OCR processing
- **tqdm**: Progress bar display
- **sumy**: Text summarization
- **keybert**: BERT-based keyword extraction
- **jieba**: Chinese word segmentation
- **transformers**: Transformer model library

## Important Notes

1. Required pre-trained models will be automatically downloaded on first run
2. PDF processing requires significant computational resources; GPU acceleration is recommended
3. Ensure input documents have correct encoding to avoid garbled text
4. For large batch processing, process in smaller batches to avoid memory overflow

## Troubleshooting

### Model Download Failure

If you encounter model download issues:

1. Set `USE_LOCAL_MODELS = False` to allow online downloads
2. Manually download models to the `models/` directory
3. Check network connection and proxy settings

### OCR Processing Failure

If PDF processing encounters problems:

1. Check CUDA environment configuration (GPU mode)
2. Set `USE_GPU_FOR_OCR = False` to use CPU mode
3. Ensure custom_models directory contains required OCR models

### Memory Shortage

If memory issues occur when processing large documents:

1. Reduce the number of documents in batch processing
2. Lower image resolution
3. Use a machine with more memory