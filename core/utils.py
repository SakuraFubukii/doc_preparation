"""
工具函数模块

包含各种辅助函数，如文本清理、元数据提取、文件处理等
"""

import re
import config
from pathlib import Path

def clean_markdown(text):
    """清理优化Markdown文本
    
    Args:
        text (str): 需要清理的Markdown文本
        
    Returns:
        str: 清理后的Markdown文本
    """
    # 合并多个连续换行为两个换行
    text = re.sub(r'\n{3,}', '\n\n', text)
    # 清理多余空格
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def extract_metadata(doc):
    """提取文档元数据
    
    Args:
        doc: Document对象，python-docx库的Document实例
        
    Returns:
        dict: 包含文档元数据的字典
    """
    cp = doc.core_properties
    attrs = ['title', 'author', 'subject', 'keywords', 'comments', 'category', 
            'content_status', 'identifier', 'language', 'version']
    metadata = {}
    
    for attr in attrs:
        metadata[attr] = getattr(cp, attr, "") or ""
    
    # 处理时间属性
    metadata["created"] = cp.created.isoformat() if cp.created else ""
    metadata["modified"] = cp.modified.isoformat() if cp.modified else ""
    metadata["revision"] = cp.revision or 0
    
    return metadata


def save_images(doc, output_folder):
    """保存文档中的图片
    
    Args:
        doc: Document对象，python-docx库的Document实例
        output_folder (str或Path): 图片保存的目标文件夹
        
    Returns:
        list: 保存的图片路径列表
    """
    saved_images = []
    try:
        images_folder = Path(output_folder) / "imgs"
        images_folder.mkdir(exist_ok=True, parents=True)
        
        for i, rel in enumerate(doc.part.rels.values()):
            if "image" in rel.target_ref:
                img_data = rel.target_part.blob
                ext = rel.target_ref.split('.')[-1].lower() if '.' in rel.target_ref else 'png'
                img_name = f"image_{i+1}.{ext}"
                img_path = images_folder / img_name
                
                with open(img_path, 'wb') as f:
                    f.write(img_data)
                saved_images.append(f"imgs/{img_name}")
    except Exception as e:
        print(f"保存图片失败: {str(e)}")
    
    return saved_images


def ensure_dir(path):
    """确保目录存在，如果不存在则创建
    
    Args:
        path (str或Path): 需要确保存在的目录路径
        
    Returns:
        Path: 目录路径对象
    """
    dir_path = Path(path)
    dir_path.mkdir(exist_ok=True, parents=True)
    return dir_path


def is_temp_file(file_path):
    """检查是否为临时文件
    
    Args:
        file_path (str或Path): 文件路径
        
    Returns:
        bool: 如果是临时文件则返回True，否则返回False
    """
    return Path(file_path).name.startswith('~$')


def is_short_text(text, threshold=20):
    """判断是否为短文本
    
    Args:
        text (str): 需要判断的文本
        threshold (int, optional): 短文本阈值，默认为20
        
    Returns:
        bool: 如果是短文本则返回True，否则返回False
    """
    return len(text.strip()) < threshold


def combine_text_fragments(fragments, separator=" "):
    """合并文本片段
    
    Args:
        fragments (list): 文本片段列表
        separator (str, optional): 分隔符，默认为空格
        
    Returns:
        str: 合并后的文本
    """
    if not fragments:
        return ""
    return separator.join(fragments)


def generate_safe_filename(text, max_length=100):
    """生成安全的文件名
    
    Args:
        text (str): 原始文本
        max_length (int, optional): 最大长度，默认为100
        
    Returns:
        str: 安全的文件名
    """
    # 移除不安全的字符
    safe_text = re.sub(r'[\\/*?:"<>|]', "", text)
    # 替换空格
    safe_text = safe_text.replace(" ", "_")
    # 限制长度
    if len(safe_text) > max_length:
        safe_text = safe_text[:max_length]
    return safe_text or "untitled"


def write_json_file(data, file_path, ensure_ascii=False, indent=2):
    """将数据写入JSON文件
    
    Args:
        data: 要写入的数据（字典或列表）
        file_path (str或Path): 文件路径
        ensure_ascii (bool, optional): 是否确保ASCII编码，默认为False
        indent (int, optional): 缩进空格数，默认为2
        
    Returns:
        bool: 写入成功返回True，否则返回False
    """
    import json
    
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)
        return True
    except Exception as e:
        print(f"写入JSON文件失败: {str(e)}")
        return False
