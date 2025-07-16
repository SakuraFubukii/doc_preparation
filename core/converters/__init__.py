"""
文档转换器模块

包含各种文档转换器的封装类，统一接口。
"""

from pathlib import Path
# 导入Word文档转换相关函数
from .docx_converter import convert_docx_to_markdown, process_single_docx

# 惰性加载PDF转换器，避免不必要的大模型加载
def get_pdf_pipeline():
    """获取PDF处理管道
    
    为了避免在不需要处理PDF时加载大模型
    """
    try:
        from .pdf_converter import get_pipeline
        return get_pipeline()
    except Exception as e:
        print(f"无法加载PDF处理模型: {str(e)}")
        return None

class DocxConverter:
    """Word 文档转换器类"""
    
    def __init__(self):
        """初始化转换器"""
        pass
        
    def convert(self, input_file, output_dir=None):
        """
        将 Word 文档转换为 Markdown
        
        Args:
            input_file: 输入文件路径
            output_dir: 输出目录路径，默认为文件名命名的子目录
            
        Returns:
            tuple: (Markdown 文本, 元数据字典)
        """
        # 确保input_file是Path对象
        input_file = Path(input_file)
        
        if output_dir is None:
            output_dir = input_file.parent / input_file.stem
        else:
            output_dir = Path(output_dir)
        
        md_content, tables_data, metadata = convert_docx_to_markdown(input_file, output_dir)
        return md_content, metadata

class PdfConverter:
    """PDF 文档转换器类"""
    
    def __init__(self):
        """初始化转换器"""
        self.pipeline = None
        try:
            # 惰性加载PDF转换管道
            self.pipeline = get_pdf_pipeline()
            if self.pipeline is None:
                print("警告: PDF处理模型加载失败，PDF转换功能将不可用")
        except Exception as e:
            print(f"初始化PDF转换器失败: {str(e)}")
        
    def convert(self, input_file, output_dir=None):
        """
        将 PDF 文档转换为 Markdown
        
        Args:
            input_file: 输入文件路径
            output_dir: 输出目录路径，默认为文件名命名的子目录
            
        Returns:
            str: Markdown 文本
        """
        # 确保input_file是Path对象
        input_file = Path(input_file)
        
        if output_dir is None:
            output_dir = input_file.parent / input_file.stem
        else:
            output_dir = Path(output_dir)
            
        # 创建输出目录
        output_dir.mkdir(exist_ok=True, parents=True)
        
        # 检查模型是否已加载
        if self.pipeline is None:
            error_msg = "PDF处理模型未加载，无法转换PDF文件"
            print(f"错误: {error_msg}")
            return f"# {input_file.stem}\n\n{error_msg}\n"
        
        try:
            from .pdf_converter import process_document
            success = process_document(str(input_file), output_dir, self.pipeline)
            
            if success:
                # 读取生成的Markdown文件
                md_file = output_dir / f"{input_file.stem}.md"
                if md_file.exists():
                    return md_file.read_text(encoding='utf-8')
                else:
                    return f"# {input_file.stem}\n\nPDF处理成功但未生成Markdown文件\n"
            else:
                return f"# {input_file.stem}\n\nPDF处理失败\n"
        except Exception as e:
            print(f"PDF转换出错: {str(e)}")
            return f"# {input_file.stem}\n\nPDF转换出错: {str(e)}\n"
