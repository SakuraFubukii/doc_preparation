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

# 确保模型缓存目录存在
os.makedirs(config.MODEL_CACHE_DIR, exist_ok=True)

# 配置模型环境
def configure_model_environment():
    """配置模型运行环境"""
    if config.USE_LOCAL_MODELS:
        os.environ["TRANSFORMERS_OFFLINE"] = "1"  # 禁止在线下载模型
        os.environ["HF_DATASETS_OFFLINE"] = "1"  # 禁止在线下载数据集
        os.environ["TRANSFORMERS_CACHE"] = str(config.MODEL_CACHE_DIR)  # 设置模型缓存目录
        os.environ["HF_HOME"] = str(config.MODEL_CACHE_DIR)  # 设置Hugging Face主目录
        
        # 检查模型是否存在
        model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
        if not model_path.exists():
            print(f"警告: 本地模型路径不存在: {model_path}")
            print(f"请确保已将模型下载到以下位置: {model_path}")
            print(f"或者设置 config.USE_LOCAL_MODELS = False 以允许从网络下载")
            return False
    return True

# 初始化模型环境
configure_model_environment()

def read_docx_content(file_path):
    """读取Word文档内容"""
    try:
        doc = Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                full_text.append(text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return ""

def generate_summary(text, sentences_count=3):
    """生成文本摘要"""
    if not text.strip():
        return []
        
    # 创建解析器
    parser = PlaintextParser.from_string(text, Tokenizer("chinese"))
    # 选择摘要算法
    summarizer = LexRankSummarizer()
    # 生成摘要
    summary = summarizer(parser.document, sentences_count=sentences_count)
    return summary

def extract_keywords(text, top_n=10):
    """从文本中提取关键词
    
    Args:
        text (str): 输入文本
        top_n (int, optional): 提取关键词的数量，默认为10
        
    Returns:
        list: 关键词列表，每项为(关键词, 权重)元组
    """
    if not text or not isinstance(text, str) or not text.strip():
        return []
    
    # 优先尝试使用jieba提取关键词
    try:
        backup_keywords = jieba.analyse.extract_tags(text, topK=top_n, withWeight=True)
        if backup_keywords:
            return backup_keywords
    except Exception as e:
        print(f"jieba.analyse提取关键词失败: {str(e)}")
    
    # 如果jieba方法失败，尝试使用KeyBERT
    try:
        # 使用jieba分词
        segmented_text = " ".join(jieba.cut(text))
        
        # 初始化KeyBERT模型
        model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
        if config.USE_LOCAL_MODELS and model_path.exists():
            kw_model = KeyBERT(model=str(model_path))
        else:
            kw_model = KeyBERT(model=config.EMBEDDING_MODEL_NAME)
        
        # 提取关键词
        keywords = kw_model.extract_keywords(segmented_text, keyphrase_ngram_range=(1, 1), top_n=top_n)
        return keywords
    
    except Exception as e:
        print(f"KeyBERT关键词提取错误: {str(e)}")
        
        # 最后尝试使用jieba的TextRank算法
        try:
            textrank_keywords = jieba.analyse.textrank(text, topK=top_n, withWeight=True)
            if textrank_keywords:
                return textrank_keywords
        except Exception as e3:
            print(f"所有关键词提取方法均失败: {str(e3)}")
        
        # 所有方法都失败，返回空列表
        return []


class MetadataExtractor:
    """元数据提取器类"""
    
    def __init__(self, summary_sentences=3, keywords_count=10):
        """初始化元数据提取器
        
        Args:
            summary_sentences: 摘要中包含的句子数
            keywords_count: 提取的关键词数量
        """
        self.summary_sentences = summary_sentences
        self.keywords_count = keywords_count
        
    def extract(self, text):
        """从文本中提取元数据
        
        Args:
            text: 输入文本
            
        Returns:
            dict: 提取的元数据字典
        """
        if not text or not isinstance(text, str):
            return {
                "summary": "",
                "keywords": {},
                "char_count": 0,
                "word_count": 0
            }
        
        try:
            # 生成摘要
            summary_sentences = generate_summary(text, self.summary_sentences)
            summary = " ".join([str(sentence) for sentence in summary_sentences])
            
            # 提取关键词
            keywords = extract_keywords(text, self.keywords_count)
            
            # 格式化关键词为字典
            keywords_dict = {keyword: float(score) for keyword, score in keywords}
            
            # 返回元数据
            return {
                "summary": summary,
                "keywords": keywords_dict,
                "char_count": len(text),
                "word_count": len(text.split())
            }
        except Exception as e:
            print(f"元数据提取错误: {str(e)}")
            # 返回基本元数据
            return {
                "summary": "",
                "keywords": {},
                "char_count": len(text),
                "word_count": len(text.split()),
                "error": str(e)
            }
        
    def extract_from_file(self, file_path):
        """从文件中提取元数据
        
        Args:
            file_path: 文件路径
            
        Returns:
            dict: 提取的元数据字典
        """
        if not Path(file_path).exists():
            return {
                "error": f"文件不存在: {file_path}",
                "summary": "",
                "keywords": {},
                "char_count": 0,
                "word_count": 0
            }
            
        # 根据文件类型选择处理方法
        if str(file_path).lower().endswith(('.docx', '.doc')):
            text = read_docx_content(file_path)
        else:
            # 对于其他文件类型，尝试作为文本文件读取
            try:
                text = Path(file_path).read_text(encoding='utf-8')
            except Exception as e:
                return {
                    "error": f"无法读取文件: {str(e)}",
                    "summary": "",
                    "keywords": {},
                    "char_count": 0,
                    "word_count": 0
                }
                
        # 提取元数据
        metadata = self.extract(text)
        metadata["file_path"] = str(file_path)
        metadata["file_name"] = Path(file_path).name
        
        return metadata


def process_documents(input_folder=None, summary_sentences=3, keywords_count=10):
    """处理所有Word文档并生成摘要和关键词
    
    Args:
        input_folder: 输入文件夹路径，默认使用config中的路径
        summary_sentences: 摘要中句子数量
        keywords_count: 提取关键词数量
        
    Returns:
        dict: 处理结果，文件名为键，内容特征为值
    """
    # 如果未指定输入文件夹，使用配置中的路径
    if input_folder is None:
        input_folder = config.INPUT_DIR
    
    # 查找所有Word文档
    word_files = list(Path(input_folder).glob("*.docx"))
    results = {}
    
    if not word_files:
        print(f"在 {input_folder} 中未找到Word文档")
        return results
    
    for file_path in word_files:
        file_name = file_path.name
        print(f"\n{'='*50}")
        print(f"处理文件: {file_name}")
        print(f"{'='*50}")
        
        # 读取文档内容
        content = read_docx_content(file_path)
        
        if not content:
            print(f"文件 {file_name} 内容为空或无法读取")
            continue
        
        # 生成摘要
        summary = generate_summary(content, summary_sentences)
        summary_texts = [str(sentence) for sentence in summary]
        
        # 提取关键词
        keywords_with_scores = extract_keywords(content, keywords_count)
        keywords_dict = {keyword: score for keyword, score in keywords_with_scores}
        
        # 创建结果字典
        result = {
            "filename": file_name,
            "keywords": keywords_dict,
            "summary": summary_texts,
            "char_count": len(content),
            "word_count": len(content.split())
        }
        
        results[file_name] = result
        
        # 为每个文件单独输出JSON格式结果
        print(f"\n文件 '{file_name}' 的JSON格式输出结果:")
        print(f"{'-'*40}")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        print(f"{'-'*40}")
    
    return results


def check_model_exists():
    """检查模型是否存在于本地缓存中"""
    if not config.USE_LOCAL_MODELS:
        print("未启用本地模型模式，将从网络下载模型")
        return True
    
    model_path = config.MODEL_CACHE_DIR / config.EMBEDDING_MODEL_NAME
    if not model_path.exists():
        print(f"警告: 本地模型路径不存在: {model_path}")
        print(f"请确保已将模型下载到以下位置: {model_path}")
        print("或者在config.py中设置 USE_LOCAL_MODELS = False 允许从网络下载")
        return False
    
    print(f"已找到本地模型: {model_path}")
    return True


# 当作为独立脚本运行时执行
if __name__ == "__main__":
    # 检查模型是否存在
    model_exists = check_model_exists()
    
    if model_exists:
        # 处理文档
        process_documents()
    else:
        print("模型检查未通过，请修复后重试")