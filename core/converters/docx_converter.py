"""
Word文档转换器模块

此模块提供Word文档(.docx)转换为Markdown文件的功能。
支持处理文本、表格和图像内容，并能提取文档元数据。
"""

import json
from docx import Document
from pathlib import Path
from tqdm import tqdm

import config
from core.utils import (
    clean_markdown, extract_metadata, save_images,
    ensure_dir, is_temp_file, is_short_text,
    combine_text_fragments, write_json_file
)

def extract_table_data(table):
    """提取表格数据"""
    headers = [cell.text.strip() for cell in table.rows[0].cells] if table.rows else []
    data_rows = []
    
    for row in table.rows[1:]:
        row_data = {}
        for i, cell in enumerate(row.cells):
            key = headers[i] if i < len(headers) else f"列{i+1}"
            row_data[key] = cell.text.strip()
        if row_data:
            data_rows.append(row_data)
    
    return headers, data_rows

def process_paragraph(para, pending_short_text):
    """处理段落文本
    
    Args:
        para: Word文档段落对象
        pending_short_text: 待处理的短文本列表
        
    Returns:
        tuple: (处理后的文本内容, 更新后的短文本列表)
    """
    text = para.text.strip()
    
    # 处理标题
    if para.style.name.startswith('Heading'):
        level = int(para.style.name.split()[-1]) if para.style.name.split()[-1].isdigit() else 2
        return f"{'#' * level} {text}\n\n", []
    
    # 处理列表
    if 'List' in para.style.name:
        prefix = '1. ' if 'Number' in para.style.name else '- '
        return f"{prefix}{text}\n", []
    
    # 处理短文本
    if is_short_text(text, config.SHORT_TEXT_THRESHOLD):
        return "", pending_short_text + [text]
    
    # 处理长文本
    return combine_text_fragments(pending_short_text + [text]) + "\n\n", []



def convert_docx_to_markdown(docx_path, output_folder):
    """将Word文档转换为Markdown格式
    
    Args:
        docx_path: Word文档的路径
        output_folder: 输出目录路径
    
    Returns:
        tuple: (Markdown内容, 表格数据列表, 元数据字典)
    """
    # 确保参数是Path对象
    docx_path = Path(docx_path)
    output_folder = Path(output_folder)
    
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"无法打开文档 {docx_path}: {str(e)}")
        return "", [], {}
    
    title = getattr(doc.core_properties, 'title', docx_path.stem)
    md_lines = [f"# {title}\n\n"]
    tables_data = []
    pending_short_text = []
    
    # 提取元数据
    metadata = extract_metadata(doc)
    
    for element in doc.element.body:
        # 处理段落
        if element.tag.endswith('p'):
            para = next((p for p in doc.paragraphs if p._element is element), None)
            if para:
                content, pending_short_text = process_paragraph(para, pending_short_text)
                if content:
                    md_lines.append(content)
        
        # 处理表格
        elif element.tag.endswith('tbl'):
            table = doc.tables[len(tables_data)]
            # 处理待处理的短文本
            if pending_short_text:
                md_lines.append(combine_text_fragments(pending_short_text) + "\n\n")
                pending_short_text = []
            
            # 表格处理：使用纯文本
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells)
                if row_text:
                    md_lines.append(f"| {row_text} |\n")
            md_lines.append("\n")
            
            # 提取表格数据
            headers, data_rows = extract_table_data(table)
            tables_data.append({"headers": headers, "data": data_rows})
    
    # 处理剩余的短文本
    if pending_short_text:
        md_lines.append(combine_text_fragments(pending_short_text) + "\n\n")
    
    # 添加图片
    for img_path in save_images(doc, output_folder):
        md_lines.append(f"![图片]({img_path})\n\n")
    
    return clean_markdown("".join(md_lines)), tables_data, metadata

def process_single_docx(input_path, output_folder):
    """处理单个Word文档
    
    Args:
        input_path: Word文档的路径
        output_folder: 输出目录路径
        
    Returns:
        bool: 处理成功返回True，否则返回False
    """
    # 确保参数是Path对象
    input_path = Path(input_path)
    output_folder = Path(output_folder)
    
    if is_temp_file(input_path):
        print(f"⚠️ 跳过临时文件: {input_path.name}")
        return False
    
    base_name = input_path.stem
    output_subfolder = ensure_dir(output_folder / base_name)
    
    try:
        md_content, tables_data, metadata = convert_docx_to_markdown(input_path, output_subfolder)
        
        # 保存Markdown文件
        (output_subfolder / f"{base_name}.md").write_text(md_content, encoding='utf-8')
        
        # 保存元数据
        write_json_file(metadata, output_subfolder / f"{base_name}_metadata.json")
        
        # 保存表格数据
        if tables_data:
            write_json_file(tables_data, output_subfolder / f"{base_name}_tables.json")
        
        print(f"✅ 成功转换: {Path(input_path).name}")
        print(f"  输出目录: {output_subfolder}")
        return True
    except Exception as e:
        print(f"❌ 处理失败: {Path(input_path).name} - {str(e)}")
        return False

def convert_batch(input_folder, output_folder, file_extensions=None):
    """批量转换Word文档
    
    Args:
        input_folder: 输入目录路径
        output_folder: 输出目录路径
        file_extensions: 支持的文件扩展名列表，默认为['.docx', '.doc']
        
    Returns:
        tuple: (成功数量, 总数量)
    """
    if file_extensions is None:
        file_extensions = ['.docx', '.doc']
    
    # 确保参数是Path对象
    input_path = Path(input_folder)
    output_folder = Path(output_folder)
    
    if not input_path.exists():
        print(f"错误: 输入文件夹 '{input_path}' 不存在")
        return 0, 0
    
    ensure_dir(output_folder)
    doc_files = [f for f in input_path.iterdir() 
                if str(f.suffix).lower() in file_extensions and not is_temp_file(f)]
    
    if not doc_files:
        print(f"在 {input_folder} 中没有找到Word文档")
        return 0, 0
    
    print(f"找到 {len(doc_files)} 个Word文档，开始转换...")
    success_count = 0
    
    for doc_file in tqdm(doc_files, desc="转换文档"):
        if process_single_docx(doc_file, output_folder):
            success_count += 1

    print("\n" + "=" * 50)
    print(f"处理完成! 成功转换: {success_count}/{len(doc_files)}")
    print(f"输出目录: {Path(output_folder).absolute()}")
    print("=" * 50)
    
    return success_count, len(doc_files)


def main():
    """作为独立模块运行时的入口点"""
    import sys
    import os
    
    # 添加项目根目录到路径
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))
    
    # 默认配置
    INPUT_FOLDER = "input"
    OUTPUT_FOLDER = "output"
    
    # 解析命令行参数
    if len(sys.argv) > 1:
        INPUT_FOLDER = sys.argv[1]
    if len(sys.argv) > 2:
        OUTPUT_FOLDER = sys.argv[2]
    
    convert_batch(INPUT_FOLDER, OUTPUT_FOLDER)


if __name__ == "__main__":
    main()