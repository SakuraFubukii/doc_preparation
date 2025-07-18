"""
文档处理主程序

处理输入目录中的Word和PDF文档，转换为Markdown并提取元数据。
该模块作为整个系统的入口点，协调各个处理组件的工作流程。
"""

import json
from pathlib import Path
from tqdm import tqdm
import config
from core.converters import DocxConverter, PdfConverter
from core.utils import ensure_dir, save_markdown_and_metadata
from core.metadata_extractor import MetadataExtractor


def save_results(markdown, metadata, output_file_path):
    """保存处理结果到指定位置
    
    将转换后的Markdown文本和提取的元数据保存到指定路径，文件名基于输出路径生成，
    元数据以JSON格式保存，支持中文字符。
    
    Args:
        markdown (str): 转换后的Markdown文本
        metadata (dict): 提取的元数据字典，包含文档属性和内容特征
        output_file_path (str): 输出文件路径（不含扩展名）
    """
    save_markdown_and_metadata(markdown, metadata, output_file_path)


def process_file(file_path, output_dir):
    """处理单个文档文件
    
    根据文件类型选择适当的转换器处理文档，提取元数据，并保存结果。
    支持Word文档(.docx/.doc)和PDF文档。
    
    Args:
        file_path (str): 文件路径
        output_dir (str): 输出目录
        
    Returns:
        bool: 处理成功返回True，否则返回False
    """
    try:
        file_path = Path(file_path)
        file_name = file_path.stem
        output_subdir = Path(output_dir) / file_name
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
            
            # 保存结果
            output_file = output_subdir / file_name
            save_results(markdown, {**metadata, **content_meta}, output_file)
            
        elif file_path_str.endswith('.pdf'):
            # PDF文档处理（现在包含清洗和标准化）
            converter = PdfConverter()
            ensure_dir(output_subdir)
            markdown, pdf_metadata = converter.convert(file_path, output_subdir)
            
            # PDF转换器已经处理了清洗、标准化和元数据提取
            if markdown and markdown.strip():
                print(f"已成功处理PDF: {file_name}")
                # 如果需要额外的内容特征提取，可以在这里添加
                if not pdf_metadata:
                    # 如果PDF转换器没有生成元数据，则手动提取
                    try:
                        extractor = MetadataExtractor(
                            summary_sentences=config.SUMMARY_SENTENCES,
                            keywords_count=config.KEYWORDS_TOP_N
                        )
                        content_meta = extractor.extract(markdown)
                        
                        # 保存元数据
                        output_file = output_subdir / file_name
                        save_results(markdown, content_meta, output_file)
                    except Exception as e:
                        print(f"处理PDF元数据失败: {str(e)}")
                        return False
            else:
                print(f"警告: PDF转换失败或生成空内容: {file_name}")
                return False
        else:
            print(f"不支持的文件类型: {file_path}")
            return False
        return True
    except Exception as e:
        print(f"处理文件失败 {file_path}: {str(e)}")
        return False


def main():
    """主函数 - 程序入口点
    
    执行完整的文档处理工作流:
    1. 验证输入输出目录
    2. 检查必要的模型资源
    3. 收集需要处理的文件
    4. 批量处理Word和PDF文档
    5. 汇总处理结果
    """
    try:
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

        # 定义支持的文件扩展名
        word_extensions = ('.docx', '.doc')
        pdf_extensions = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff')
        
        # 获取所有文件
        word_files = [f for ext in word_extensions for f in input_path.glob(f"**/*{ext}")]
        pdf_files = [f for ext in pdf_extensions for f in input_path.glob(f"**/*{ext}")]
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
            from core.converters.pdf_converter import process_directory
            pdf_success_count, _ = process_directory(input_dir, output_dir)
            success_count += pdf_success_count

        print("\n" + "=" * 50)
        print(f"处理完成! 成功: {success_count}/{len(all_files)}")
        print(f"输出目录: {Path(output_dir).absolute()}")
        print("=" * 50)
        
    except Exception as e:
        print(f"程序执行过程中出现错误: {str(e)}")
    finally:
        # 清理全局资源
        try:
            print("正在清理系统资源...")
            from core.converters.pdf_converter import cleanup_resources
            cleanup_resources()
            
            # 强制垃圾回收
            import gc
            gc.collect()
            print("系统资源清理完成")
        except Exception as e:
            print(f"清理资源时出错: {str(e)}")


if __name__ == "__main__":
    main()
