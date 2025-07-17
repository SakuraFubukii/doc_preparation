"""
配置文件

包含文档处理系统所需的全局配置参数，包括:
- 目录路径配置
- 文档处理参数
- 预训练模型设置
- PDF处理选项
"""

import os
from pathlib import Path

# 基础目录配置
# BASE_DIR: 项目根目录的绝对路径
# INPUT_DIR: 待处理文档的输入目录
# OUTPUT_DIR: 处理结果的输出目录
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = r"E:\Document\petrochina\知识问答工作流平台\input"
OUTPUT_DIR = r"E:\Document\petrochina\知识问答工作流平台\output"

# 文档处理相关配置
# SUMMARY_SENTENCES: 生成摘要的句子数量
# KEYWORDS_TOP_N: 提取关键词的数量
# SHORT_TEXT_THRESHOLD: 判断为短文本的字符阈值
SUMMARY_SENTENCES = 3
KEYWORDS_TOP_N = 10
SHORT_TEXT_THRESHOLD = 20

# 预训练模型相关配置
# MODEL_CACHE_DIR: 模型缓存目录路径
# EMBEDDING_MODEL_NAME: 使用的嵌入模型名称
# USE_LOCAL_MODELS: 是否使用本地模型（不从网络下载）
MODEL_CACHE_DIR = BASE_DIR / "models"
EMBEDDING_MODEL_NAME = "all-MiniLM-L6-v2"
USE_LOCAL_MODELS = True

# PDF处理相关配置
# USE_GPU_FOR_OCR: 是否使用GPU加速OCR处理
USE_GPU_FOR_OCR = True