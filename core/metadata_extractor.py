"""
元数据提取模块

从文本中提取关键信息，如摘要、关键词等
"""

import os
import json
from pathlib import Path
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
from keybert import KeyBERT
import jieba
import jieba.analyse
import config
from docx import Document

# 确保模型缓存目录存在
os.makedirs(config.MODEL_CACHE_DIR, exist_ok=True)

def configure_model_environment():
    """配置模型运行环境"""
    if config.USE_LOCAL_MODELS:
        os.environ.update({
            "TRANSFORMERS_OFFLINE": "1",
            "HF_DATASETS_OFFLINE": "1",
            "TRANSFORMERS_CACHE": str(config.MODEL_CACHE_DIR),
            "HF_HOME": str(config.MODEL_CACHE_DIR)
        })
        
        # 检查模型是否存在
        model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
        if not model_path.exists():
            print(f"警告: 本地模型路径不存在: {model_path}")
            return False
    return True

# 初始化模型环境
configure_model_environment()

def read_docx_content(file_path):
    """读取Word文档内容"""
    try:
        doc = Document(file_path)
        return "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return ""

def generate_summary(text, sentences_count=3):
    """生成文本摘要"""
    if not text.strip():
        return []
    
    parser = PlaintextParser.from_string(text, Tokenizer("chinese"))
    return LexRankSummarizer()(parser.document, sentences_count=sentences_count)

def extract_keywords(text, top_n=10):
    """从文本中提取关键词"""
    if not text or not isinstance(text, str) or not text.strip():
        return []
    
    # 尝试使用jieba提取关键词
    try:
        keywords = jieba.analyse.extract_tags(text, topK=top_n, withWeight=True)
        if keywords:
            return keywords
    except Exception as e:
        print(f"jieba关键词提取失败: {str(e)}")
    
    # 备选方案：使用KeyBERT
    try:
        segmented_text = " ".join(jieba.cut(text))
        model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
        model_source = str(model_path) if config.USE_LOCAL_MODELS and model_path.exists() else config.EMBEDDING_MODEL_NAME
        keywords = KeyBERT(model=model_source).extract_keywords(segmented_text, keyphrase_ngram_range=(1, 1), top_n=top_n)
        return keywords
    except Exception as e:
        print(f"KeyBERT关键词提取错误: {str(e)}")
    
    # 最后尝试使用TextRank
    try:
        return jieba.analyse.textrank(text, topK=top_n, withWeight=True)
    except Exception as e:
        print(f"关键词提取失败: {str(e)}")
        return []


class MetadataExtractor:
    """元数据提取器类"""
    
    def __init__(self, summary_sentences=3, keywords_count=10):
        self.summary_sentences = summary_sentences
        self.keywords_count = keywords_count
        
    def extract(self, text):
        """从文本中提取元数据"""
        if not text or not isinstance(text, str):
            return {"summary": "", "keywords": {}, "char_count": 0, "word_count": 0}
        
        try:
            # 生成摘要
            summary_sentences = generate_summary(text, self.summary_sentences)
            summary = " ".join([str(sentence) for sentence in summary_sentences])
            
            # 提取关键词
            keywords = extract_keywords(text, self.keywords_count)
            keywords_dict = {keyword: float(score) for keyword, score in keywords}
            
            return {
                "summary": summary,
                "keywords": keywords_dict,
                "char_count": len(text),
                "word_count": len(text.split())
            }
        except Exception as e:
            print(f"元数据提取错误: {str(e)}")
            return {
                "summary": "",
                "keywords": {},
                "char_count": len(text),
                "word_count": len(text.split()),
                "error": str(e)
            }
        
    def extract_from_file(self, file_path):
        """从文件中提取元数据"""
        file_path = Path(file_path)
        if not file_path.exists():
            return {"error": f"文件不存在: {file_path}", "summary": "", "keywords": {}, 
                    "char_count": 0, "word_count": 0}
        
        try:
            # 根据文件类型选择处理方法
            if file_path.suffix.lower() in ('.docx', '.doc'):
                text = read_docx_content(file_path)
            else:
                text = file_path.read_text(encoding='utf-8')
                
            metadata = self.extract(text)
            metadata.update({"file_path": str(file_path), "file_name": file_path.name})
            return metadata
            
        except Exception as e:
            return {"error": f"无法读取文件: {str(e)}", "summary": "", "keywords": {},
                    "char_count": 0, "word_count": 0}


def process_documents_batch(input_folder=None, summary_sentences=3, keywords_count=10):
    """批量处理所有Word文档并生成摘要和关键词"""
    input_folder = input_folder or config.INPUT_DIR
    word_files = list(Path(input_folder).glob("*.docx"))
    results = {}
    
    if not word_files:
        print(f"在 {input_folder} 中未找到Word文档")
        return results
    
    extractor = MetadataExtractor(summary_sentences, keywords_count)
    
    for file_path in word_files:
        file_name = file_path.name
        print(f"\n{'='*50}\n处理文件: {file_name}\n{'='*50}")
        
        try:
            result = extractor.extract_from_file(file_path)
            results[file_name] = result
            print(f"\n文件 '{file_name}' 处理结果:\n{'-'*40}")
            print(json.dumps(result, ensure_ascii=False, indent=2))
            print(f"{'-'*40}")
        except Exception as e:
            print(f"处理文件 {file_name} 失败: {str(e)}")
            results[file_name] = {"error": str(e)}
    
    return results


def check_model_exists():
    """检查模型是否存在于本地缓存中"""
    if not config.USE_LOCAL_MODELS:
        print("未启用本地模型模式，将从网络下载模型")
        return True
    
    model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
    if not model_path.exists():
        print(f"警告: 本地模型路径不存在: {model_path}")
        return False
    
    print(f"已找到本地模型: {model_path}")
    return True


if __name__ == "__main__":
    if check_model_exists():
        process_documents_batch()
    else:
        print("模型检查未通过，请修复后重试")