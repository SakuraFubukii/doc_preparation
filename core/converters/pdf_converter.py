"""
PDF文档转换器模块

此模块提供PDF文档转换为Markdown文件的功能，使用PaddleOCR进行处理。
"""

import time
import json
from pathlib import Path
from paddleocr import PPStructureV3
import os
import shutil
import config

# 设置设备类型
device = 'gpu' if config.USE_GPU_FOR_OCR else 'cpu'

# 初始化PaddleOCR结构化文档分析管线
pipeline = PPStructureV3(
    device=device,
)

print("OCR模型加载完成")

supported_extensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']

def process_document(input_file, output_root):
    file_path = Path(input_file)
    if file_path.suffix.lower() not in supported_extensions:
        raise ValueError(f"不支持的文类型: {file_path.suffix}")
    
    start_time = time.time()
    print(f"\n开始处理: {file_path.name}")
    
    # 定义变量，以便在finally块中可以访问
    output_folder = None
    
    try:
        # 确保文件存在
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
            
        # 为每个文件创建独立的输出文件夹
        output_folder = Path(output_root) / file_path.stem
        output_folder.mkdir(parents=True, exist_ok=True)
        
        # 创建img子文件夹
        img_folder = output_folder / "imgs"
        img_folder.mkdir(exist_ok=True, parents=True)

        # 处理文件
        output = pipeline.predict(str(file_path))

        markdown_list = []
        markdown_images = []

        for res in output:
            md_info = res.markdown
            markdown_list.append(md_info)
            markdown_images.append(md_info.get("markdown_images", {}))
    
        markdown_texts = pipeline.concatenate_markdown_pages(markdown_list)

        # 输出Markdown文件
        mkd_file_path = output_folder / f"{file_path.stem}.md"
        with open(mkd_file_path, "w", encoding="utf-8") as f:
            f.write(markdown_texts)

        # 输出图片到img子文件夹
        for item in markdown_images:
            if item:
                for path, image in item.items():
                    # 只保留文件名，忽略原始路径
                    filename = Path(path).name
                    img_path = img_folder / filename
                    image.save(img_path)
        
        # 生成元数据
        try:
            # 动态导入MetadataExtractor，避免循环导入
            from core.metadata_extractor import MetadataExtractor
            extractor = MetadataExtractor(
                summary_sentences=config.SUMMARY_SENTENCES,
                keywords_count=config.KEYWORDS_TOP_N
            )
            # 读取生成的Markdown内容
            markdown_content = mkd_file_path.read_text(encoding='utf-8')
            # 提取元数据
            content_meta = extractor.extract(markdown_content)
            
            # 保存元数据
            meta_file = output_folder / f"{file_path.stem}_metadata.json"
            with open(meta_file, 'w', encoding='utf-8') as f:
                json.dump(content_meta, f, ensure_ascii=False, indent=2)
                
            print(f"  - 已生成元数据: {meta_file.name}")
        except Exception as meta_error:
            print(f"  - 元数据生成失败: {str(meta_error)}")
                    
        print(f"完成处理: {file_path.name}")
        print(f"  - 耗时: {time.time()-start_time:.2f}秒")
        print(f"  - 输出位置: {output_folder}")
        
        return True
    except Exception as e:
        print(f"处理文件 {file_path.name} 时出错: {str(e)}")
        # 清理部分生成的文件
        if output_folder and output_folder.exists():
            try:
                shutil.rmtree(output_folder)
            except Exception as clean_error:
                print(f"清理输出目录失败: {str(clean_error)}")
        return False

def process_directory(input_root, output_root):
    """
    处理指定目录下的所有支持的文档（PDF及图片）。

    Args:
        input_root (str): 输入目录的路径。
        output_root (str): 输出目录的路径。

    Returns:
        tuple: (成功处理的文件数, 找到的总文件数)
    """
    # 确保输入输出目录存在
    Path(input_root).mkdir(parents=True, exist_ok=True)
    Path(output_root).mkdir(parents=True, exist_ok=True)
    
    # 遍历所有支持的文件
    total_files = 0
    processed_files = 0
    failed_files = []
    
    all_files_to_process = []
    for ext in supported_extensions:
        all_files_to_process.extend(Path(input_root).glob(f"**/*{ext}"))

    total_files = len(all_files_to_process)

    if total_files > 0:
        for file_path in all_files_to_process:
            success = process_document(str(file_path), output_root)
            if success:
                processed_files += 1
            else:
                failed_files.append(str(file_path))
    
    print(f"\nPDF及图片处理完成! 共找到 {total_files} 个文件")
    print(f"  - 成功处理: {processed_files} 个文件")
    
    if failed_files:
        print(f"  - 处理失败: {len(failed_files)} 个文件")
        for f in failed_files:
            print(f"    - {f}")
            
    return processed_files, total_files

if __name__ == "__main__":
    input_root = r"E:\Document\petrochina\知识问答工作流平台\doc_preparation\input"
    output_root = r"E:\Document\petrochina\知识问答工作流平台\doc_preparation\output"
    process_directory(input_root, output_root)
