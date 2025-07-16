"""
文档处理主程序

处理输入目录中的Word和PDF文档，转换为Markdown并提取元数据
"""

import json
from pathlib import Path
from tqdm import tqdm
import config
from core.converters import DocxConverter, PdfConverter
from core.utils import ensure_dir
from core.metadata_extractor import MetadataExtractor


# MetadataExtractor 类已经在 core.metadata_extractor 模块中定义，直接导入使用


def save_results(markdown, metadata, output_file_path):
    """保存处理结果
    
    Args:
        markdown: Markdown文本
        metadata: 元数据字典
        output_file_path: 输出文件路径（不含扩展名）
    """
    # 创建输出目录
    output_dir = Path(output_file_path).parent
    ensure_dir(output_dir)
    
    # 保存Markdown文件
    md_file = f"{output_file_path}.md"
    with open(md_file, 'w', encoding='utf-8') as f:
        f.write(markdown)
    
    # 保存元数据文件
    meta_file = f"{output_file_path}_metadata.json"
    with open(meta_file, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    
    print(f"已保存: {Path(md_file).name}")


def process_file(file_path, output_dir):
    """处理单个文件
    
    Args:
        file_path: 文件路径
        output_dir: 输出目录
        
    Returns:
        bool: 处理成功返回True，否则返回False
    """
    try:
        # 确保file_path是Path对象
        file_path = Path(file_path)
        file_name = file_path.stem
        output_subdir = Path(output_dir) / file_name
        
        # 转换Path对象为字符串后再调用lower()
        file_path_str = str(file_path).lower()
        
        if file_path_str.endswith(('.docx', '.doc')):
            # Word文档处理
            converter = DocxConverter()
            markdown, metadata = converter.convert(file_path, output_subdir)
            
            # 提取内容特征
            extractor = MetadataExtractor(
                summary_sentences=config.SUMMARY_SENTENCES,
                keywords_count=config.KEYWORDS_TOP_N
            )
            content_meta = extractor.extract(markdown)
            
            # 合并元数据
            full_metadata = {**metadata, **content_meta}
            
            # 保存结果
            output_file = output_subdir / file_name
            save_results(markdown, full_metadata, output_file)
            
        elif file_path_str.endswith(('.pdf')):
            # PDF文档处理 - 先通过OCR生成Markdown文件
            converter = PdfConverter()
            
            # 确保输出目录存在
            ensure_dir(output_subdir)
            
            # 获取converter.convert的返回值(markdown文本)
            markdown = converter.convert(file_path, output_subdir)
            
            # 正确构建markdown文件路径
            # 一般情况下，markdown文件应该位于 output_subdir/file_name.md
            md_file = output_subdir / f"{file_name}.md"
            
            # 如果直接路径不存在，尝试检查嵌套路径
            if not md_file.exists():
                nested_md_file = output_subdir / file_name / f"{file_name}.md"
                if nested_md_file.exists():
                    # 使用嵌套路径的文件
                    md_file = nested_md_file
                    print(f"找到嵌套目录中的Markdown文件: {md_file}")
            
            # 确认Markdown文件已生成
            if md_file.exists():
                try:
                    # 读取生成的Markdown文件
                    markdown = md_file.read_text(encoding='utf-8')
                    
                    # 提取内容特征
                    extractor = MetadataExtractor(
                        summary_sentences=config.SUMMARY_SENTENCES,
                        keywords_count=config.KEYWORDS_TOP_N
                    )
                    content_meta = extractor.extract(markdown)
                    
                    # 保存元数据 - 保存到与markdown文件相同的目录
                    meta_file = md_file.parent / f"{file_name}_metadata.json"
                    with open(meta_file, 'w', encoding='utf-8') as f:
                        json.dump(content_meta, f, ensure_ascii=False, indent=2)
                    
                    print(f"已成功处理PDF并生成元数据: {file_name}")
                except Exception as e:
                    print(f"处理PDF元数据失败: {str(e)}")
                    return False
            else:
                print(f"警告: OCR处理后的Markdown文件未生成，尝试过以下路径:")
                print(f"1. {md_file}")
                print(f"2. {output_subdir / file_name / f'{file_name}.md'}")
                return False
            
        else:
            print(f"不支持的文件类型: {file_path}")
            return False
        return True
    except Exception as e:
        print(f"处理文件失败 {file_path}: {str(e)}")
        return False


def main():
    """主函数"""
    input_dir = config.INPUT_DIR
    output_dir = config.OUTPUT_DIR
    
    # 确保输入输出目录存在
    input_path = Path(input_dir)
    if not input_path.exists():
        print(f"错误: 输入目录不存在 - {input_dir}")
        return
    
    ensure_dir(output_dir)
    
    # 检查元数据提取相关模型
    try:
        from core.metadata_extractor import check_model_exists
        check_model_exists()
    except Exception as e:
        print(f"模型检查失败: {str(e)}")
        print("程序将继续运行，但元数据提取功能可能受限")

    # 获取所有支持的文件
    word_extensions = ('.docx', '.doc')
    pdf_extensions = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff')
    
    # 获取所有文件（Word文件优先）
    word_files = [f for ext in word_extensions for f in input_path.glob(f"**/*{ext}")]
    pdf_files = [f for ext in pdf_extensions for f in input_path.glob(f"**/*{ext}")]
    
    # 合并文件列表，Word文件在前
    all_files = word_files + pdf_files
    
    if not all_files:
        print(f"在 {input_dir} 中没有找到支持的文档文件")
        return
    
    print(f"找到 {len(all_files)} 个文件 (Word: {len(word_files)}, PDF及图片: {len(pdf_files)})，开始处理...")
    success_count = 0
    
    # 处理Word文件
    if word_files:
        print("\n--- 开始处理Word文件 ---")
        for file_path in tqdm(word_files, desc="处理Word文件"):
            if process_file(file_path, output_dir):
                success_count += 1

    # 处理PDF和图片文件
    if pdf_files:
        print("\n--- 开始处理PDF及图片文件 ---")
        # 对于PDF和图片，我们直接调用其自身的处理流程
        # 因为它们通常是批量OCR处理，而不是逐个提取元数据
        from core.converters.pdf_converter import process_directory
        pdf_success_count, pdf_total_count = process_directory(input_dir, output_dir)
        success_count += pdf_success_count

    print("\n" + "=" * 50)
    print(f"处理完成! 成功: {success_count}/{len(all_files)}")
    print(f"输出目录: {Path(output_dir).absolute()}")
    print("=" * 50)


if __name__ == "__main__":
    main()
