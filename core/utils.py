"""
工具函数模块

包含各种辅助函数，如文本清理、元数据提取、文件处理等
"""

import re
import config
from pathlib import Path

def clean_markdown(text):
    """清理优化Markdown文本"""
    # 合并多个连续换行为两个换行
    text = re.sub(r'\n{3,}', '\n\n', text)
    # 清理多余空格
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def extract_metadata(doc):
    """提取文档元数据"""
    cp = doc.core_properties
    attrs = ['title', 'author', 'subject', 'keywords', 'comments', 'category', 
            'content_status', 'identifier', 'language', 'version']
    
    # 构建基本元数据
    metadata = {attr: getattr(cp, attr, "") or "" for attr in attrs}
    
    # 处理时间属性
    metadata.update({
        "created": cp.created.isoformat() if cp.created else "",
        "modified": cp.modified.isoformat() if cp.modified else "",
        "revision": cp.revision or 0
    })
    
    return metadata


def save_images(doc, output_folder):
    """保存文档中的图片"""
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
    """确保目录存在，如果不存在则创建"""
    dir_path = Path(path)
    dir_path.mkdir(exist_ok=True, parents=True)
    return dir_path


def is_temp_file(file_path):
    """检查是否为临时文件"""
    return Path(file_path).name.startswith('~$')


def is_short_text(text, threshold=20):
    """判断是否为短文本"""
    return len(text.strip()) < threshold


def combine_text_fragments(fragments, separator=" "):
    """合并文本片段"""
    return separator.join(fragments) if fragments else ""


def generate_safe_filename(text, max_length=100):
    """生成安全的文件名"""
    safe_text = re.sub(r'[\\/*?:"<>|]', "", text).replace(" ", "_")
    return (safe_text[:max_length] if len(safe_text) > max_length else safe_text) or "untitled"


def write_json_file(data, file_path, ensure_ascii=False, indent=2):
    """将数据写入JSON文件"""
    import json
    
    try:
        ensure_dir(Path(file_path).parent)
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)
        return True
    except Exception as e:
        print(f"写入JSON文件失败: {str(e)}")
        return False


def save_markdown_and_metadata(markdown_content, metadata, output_path):
    """保存Markdown内容和元数据到指定路径
    
    Args:
        markdown_content (str): Markdown文本内容
        metadata (dict): 元数据字典
        output_path (str|Path): 输出文件路径（不含扩展名）
    
    Returns:
        tuple: (md_file_path, meta_file_path) 保存的文件路径
    """
    output_path = Path(output_path)
    ensure_dir(output_path.parent)
    
    # 保存Markdown文件
    md_file = f"{output_path}.md"
    try:
        with open(md_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
    except Exception as e:
        print(f"保存Markdown文件失败: {str(e)}")
        return None, None
    
    # 保存元数据文件
    meta_file = f"{output_path}_metadata.json"
    if write_json_file(metadata, meta_file):
        print(f"已保存: {Path(md_file).name}")
        return md_file, meta_file
    else:
        return md_file, None
