"""
配置文件

包含文档处理系统所需的全局配置参数
"""

import os
from pathlib import Path

# 基础目录配置
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = "E:\Document\petrochina\知识问答工作流平台\input"  # 输入文档目录
OUTPUT_DIR = "E:\Document\petrochina\知识问答工作流平台\output"  # 处理结果输出目录

# 文档处理相关配置
SUMMARY_SENTENCES = 3  # 摘要句子数量
KEYWORDS_TOP_N = 10    # 关键词提取数量
SHORT_TEXT_THRESHOLD = 20  # 短文本阈值，合并短于此长度的文本片段

# 预训练模型相关配置
MODEL_CACHE_DIR = BASE_DIR / "models"  # 预训练模型缓存目录
EMBEDDING_MODEL_NAME = "all-MiniLM-L6-v2"  # 词嵌入模型名称
USE_LOCAL_MODELS = True  # 是否使用本地模型

# PDF处理相关配置
USE_GPU_FOR_OCR = True  # 是否使用GPU进行OCR处理