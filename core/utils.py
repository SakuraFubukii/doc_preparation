"""
工具函数模块

包含各种辅助函数，如文本清理、元数据提取、文件处理等
"""

import re
import json
from pathlib import Path

def clean_markdown(text):
    """清理优化Markdown文本"""
    # 合并多个连续换行为两个换行
    text = re.sub(r'\n{3,}', '\n\n', text)
    # 清理多余空格
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def clean_ocr_text(text):
    """清洗和标准化OCR识别的文本
    
    适用于PDF OCR识别后的文本清洗，也可用于其他需要文本标准化的场景。
    注意：不包含OCR错误修复功能，避免误改正确内容。
    
    Args:
        text (str): 原始文本
        
    Returns:
        str: 清洗后的文本
    """
    if not text or not text.strip():
        return ""
    
    # 1. 基础清洗：移除多余空白字符
    text = re.sub(r'\s+', ' ', text.strip())
    
    # 2. 修复中文标点符号
    chinese_punctuation_fixes = {
        r'\s+，': '，',  # 中文逗号前后空格
        r'，\s+': '，',
        r'\s+。': '。',  # 中文句号前后空格
        r'。\s+': '。',
        r'\s+；': '；',  # 中文分号前后空格
        r'；\s+': '；',
        r'\s+：': '：',  # 中文冒号前后空格
        r'：\s+': '：',
        r'\s+！': '！',  # 中文感叹号前后空格
        r'！\s+': '！',
        r'\s+？': '？',  # 中文问号前后空格
        r'？\s+': '？',
        r'"\s+': '"',   # 中文引号处理
        r'\s+"': '"',
    }
    
    for pattern, replacement in chinese_punctuation_fixes.items():
        text = re.sub(pattern, replacement, text)
    
    # 3. 处理数字和单位之间的空格
    text = re.sub(r'(\d+)\s+([%°℃℉])', r'\1\2', text)  # 移除数字和单位符号间的空格
    text = re.sub(r'(\d+)\s+(万|千|百|十|亿|元|米|公里|公斤|吨)', r'\1\2', text)  # 中文数量单位
    
    return text.strip()


def clean_table_line(table_line):
    """清洗表格行
    
    Args:
        table_line (str): 表格行文本
        
    Returns:
        str: 清洗后的表格行
    """
    if not table_line or not table_line.strip():
        return ""
    
    # 分割表格单元格
    cells = table_line.split('|')[1:-1]  # 去掉首尾的空字符串
    
    # 清洗每个单元格
    cleaned_cells = []
    for cell in cells:
        cleaned_cell = clean_ocr_text(cell.strip())
        cleaned_cells.append(cleaned_cell)
    
    # 重新组装表格行
    if any(cell.strip() for cell in cleaned_cells):  # 确保至少有一个非空单元格
        return '| ' + ' | '.join(cleaned_cells) + ' |'
    
    return ""


def normalize_markdown_structure(markdown_text):
    """标准化Markdown文档结构
    
    适用于PDF OCR和Word转换后的Markdown文本结构化处理。
    
    Args:
        markdown_text (str): 原始Markdown文本
        
    Returns:
        str: 标准化后的Markdown文本
    """
    if not markdown_text:
        return ""
    
    lines = markdown_text.split('\n')
    processed_lines = []
    current_paragraph = []
    
    for line in lines:
        line = line.strip()
        
        # 处理标题行
        if line.startswith('#'):
            # 先保存当前段落
            if current_paragraph:
                combined_text = combine_text_fragments(current_paragraph)
                if combined_text:
                    processed_lines.append(combined_text)
                current_paragraph = []
            
            # 添加标题，确保标题格式正确
            level = len(line) - len(line.lstrip('#'))
            title_text = line.lstrip('#').strip()
            if title_text:
                processed_lines.append(f"{'#' * min(level, 6)} {title_text}")
            continue
        
        # 处理表格行
        if line.startswith('|') and line.endswith('|'):
            # 先保存当前段落
            if current_paragraph:
                combined_text = combine_text_fragments(current_paragraph)
                if combined_text:
                    processed_lines.append(combined_text)
                current_paragraph = []
            
            # 清洗表格行
            cleaned_table_line = clean_table_line(line)
            if cleaned_table_line:
                processed_lines.append(cleaned_table_line)
            continue
        
        # 处理列表项
        if re.match(r'^[-*+]\s+|^\d+\.\s+', line):
            # 先保存当前段落
            if current_paragraph:
                combined_text = combine_text_fragments(current_paragraph)
                if combined_text:
                    processed_lines.append(combined_text)
                current_paragraph = []
            
            processed_lines.append(line)
            continue
        
        # 处理空行
        if not line:
            if current_paragraph:
                combined_text = combine_text_fragments(current_paragraph)
                if combined_text:
                    processed_lines.append(combined_text)
                current_paragraph = []
            continue
        
        # 累积普通文本行
        cleaned_line = clean_ocr_text(line)
        if cleaned_line:
            current_paragraph.append(cleaned_line)
    
    # 处理最后的段落
    if current_paragraph:
        combined_text = combine_text_fragments(current_paragraph)
        if combined_text:
            processed_lines.append(combined_text)
    
    # 组装最终文本，确保适当的行间距
    result_lines = []
    for i, line in enumerate(processed_lines):
        result_lines.append(line)
        
        # 在标题、段落、表格、列表之间添加适当的空行
        if i < len(processed_lines) - 1:
            next_line = processed_lines[i + 1]
            if (line.startswith('#') or 
                next_line.startswith('#') or 
                next_line.startswith('|') or 
                re.match(r'^[-*+]\s+|^\d+\.\s+', next_line)):
                result_lines.append('')
    
    return '\n'.join(result_lines)


def post_process_markdown_content(markdown_content):
    """对转换后的Markdown内容进行完整的后处理
    
    综合应用所有清洗和标准化功能，适用于PDF和Word转换后的文本。
    
    Args:
        markdown_content (str): 原始Markdown内容
        
    Returns:
        str: 处理后的Markdown内容
    """
    # 1. 标准化结构
    markdown_content = normalize_markdown_structure(markdown_content)
    
    # 2. 应用基础清洗
    markdown_content = clean_markdown(markdown_content)
    
    # 3. 最终整理：确保文档结构清晰
    lines = markdown_content.split('\n')
    final_lines = []
    
    for i, line in enumerate(lines):
        if line.strip():
            final_lines.append(line)
        elif i > 0 and i < len(lines) - 1:
            # 保留必要的空行，但避免连续多个空行
            if final_lines and final_lines[-1].strip():
                final_lines.append('')
    
    return '\n'.join(final_lines).strip()


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


def extract_tables_from_markdown_and_save_json(md_file_path, output_folder=None):
    """从Markdown文件中提取表格并保存为JSON格式的txt文件
    
    Args:
        md_file_path (str|Path): Markdown文件路径
        output_folder (str|Path, optional): 输出文件夹，默认为md文件所在目录
    
    Returns:
        list: 生成的表格JSON文件路径列表
    """
    md_file_path = Path(md_file_path)
    if output_folder is None:
        output_folder = md_file_path.parent
    else:
        output_folder = Path(output_folder)
    
    # 读取Markdown文件
    try:
        with open(md_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"读取Markdown文件失败: {str(e)}")
        return []
    
    # 使用正则表达式匹配表格
    table_pattern = r'\|.*\|(?:\n\|.*\|)*'
    tables = re.findall(table_pattern, content, re.MULTILINE)
    
    generated_files = []
    file_base_name = md_file_path.stem
    
    for i, table_text in enumerate(tables):
        if not table_text.strip():
            continue
            
        # 解析表格
        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
        if not lines:
            continue
        
        # 提取表头和数据
        headers = []
        data_rows = []
        
        for j, line in enumerate(lines):
            # 清理表格行，移除首尾的|符号
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            
            if j == 0:
                headers = cells
            else:
                # 跳过分隔行（通常包含 --- 这样的内容）
                if all(cell.replace('-', '').replace(' ', '') == '' for cell in cells):
                    continue
                    
                if cells and len(cells) > 0:
                    row_data = {}
                    for k, cell in enumerate(cells):
                        key = headers[k] if k < len(headers) else f"列{k+1}"
                        row_data[key] = cell
                    if any(value.strip() for value in row_data.values()):  # 确保行不为空
                        data_rows.append(row_data)
        
        # 只有当表格有有效数据时才保存
        if headers and data_rows:
            table_data = {
                "headers": headers,
                "data": data_rows
            }
            
            # 生成文件名
            table_filename = f"{file_base_name}_table_{i + 1}.txt"
            table_file_path = output_folder / table_filename
            
            # 保存为JSON格式的txt文件
            try:
                with open(table_file_path, 'w', encoding='utf-8') as f:
                    json.dump(table_data, f, ensure_ascii=False, indent=2)
                print(f"  - 已生成表格文件: {table_filename}")
                generated_files.append(table_file_path)
            except Exception as e:
                print(f"  - 保存表格文件失败: {str(e)}")
    
    return generated_files
