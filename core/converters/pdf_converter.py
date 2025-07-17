"""
PDF文档转换器模块

此模块提供PDF文档转换为Markdown文件的功能，使用PaddleOCR进行处理。
"""

import time
import json
from pathlib import Path
import os
import shutil
import config
import gc
import atexit
import tempfile
import uuid

# 全局变量存储pipeline实例
_pipeline = None
_pipeline_initialized = False

supported_extensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff']

def get_pipeline():
    """获取或初始化PaddleOCR pipeline"""
    global _pipeline, _pipeline_initialized
    
    if not _pipeline_initialized:
        try:
            from paddleocr import PPStructureV3
            device = 'gpu' if config.USE_GPU_FOR_OCR else 'cpu'
            _pipeline = PPStructureV3(device=device)
            _pipeline_initialized = True
            print("OCR模型加载完成")
            
            # 注册程序退出时的清理函数
            atexit.register(cleanup_resources)
        except Exception as e:
            print(f"初始化OCR模型失败: {str(e)}")
            _pipeline = None
            _pipeline_initialized = True
    
    return _pipeline

def process_document(input_file, output_root):
    file_path = Path(input_file)
    if file_path.suffix.lower() not in supported_extensions:
        raise ValueError(f"不支持的文件类型: {file_path.suffix}")
    
    start_time = time.time()
    print(f"\n开始处理: {file_path.name}")
    output_folder = None
    temp_file_path = None
    
    try:
        # 准备目录
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        output_folder = Path(output_root) / file_path.stem
        output_folder.mkdir(parents=True, exist_ok=True)
        img_folder = output_folder / "imgs"
        img_folder.mkdir(exist_ok=True, parents=True)

        # 获取pipeline实例
        pipeline = get_pipeline()
        if pipeline is None:
            raise RuntimeError("OCR模型未能正确初始化")

        # 处理中文路径问题：如果路径包含非ASCII字符，则复制到临时文件
        processing_file_path = str(file_path)
        try:
            # 检查路径是否包含非ASCII字符
            str(file_path).encode('ascii')
        except UnicodeEncodeError:
            # 路径包含中文或其他非ASCII字符，创建临时文件
            print("  - 检测到中文路径，创建临时文件...")
            temp_dir = tempfile.gettempdir()
            temp_filename = f"temp_{uuid.uuid4().hex}{file_path.suffix}"
            temp_file_path = Path(temp_dir) / temp_filename
            shutil.copy2(file_path, temp_file_path)
            processing_file_path = str(temp_file_path)
            print(f"  - 临时文件: {temp_file_path}")

        # 处理文件
        output = pipeline.predict(processing_file_path)
        markdown_list = []
        markdown_images = []

        for res in output:
            md_info = res.markdown
            markdown_list.append(md_info)
            markdown_images.append(md_info.get("markdown_images", {}))
    
        markdown_texts = pipeline.concatenate_markdown_pages(markdown_list)

        # 保存输出
        mkd_file_path = output_folder / f"{file_path.stem}.md"
        with open(mkd_file_path, "w", encoding="utf-8") as f:
            f.write(markdown_texts)

        # 保存图片
        for item in markdown_images:
            if item:
                for path, image in item.items():
                    filename = Path(path).name
                    img_path = img_folder / filename
                    image.save(img_path)
        
        # 生成元数据
        try:
            from core.metadata_extractor import MetadataExtractor
            extractor = MetadataExtractor(
                summary_sentences=config.SUMMARY_SENTENCES,
                keywords_count=config.KEYWORDS_TOP_N
            )
            markdown_content = mkd_file_path.read_text(encoding='utf-8')
            content_meta = extractor.extract(markdown_content)
            
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
        if output_folder and output_folder.exists():
            try:
                shutil.rmtree(output_folder)
            except Exception as clean_error:
                print(f"清理输出目录失败: {str(clean_error)}")
        return False
    finally:
        # 清理临时文件
        if temp_file_path and temp_file_path.exists():
            try:
                temp_file_path.unlink()
                print("  - 已清理临时文件")
            except Exception as temp_error:
                print(f"  - 清理临时文件失败: {str(temp_error)}")

def process_directory(input_root, output_root):
    """处理指定目录下的所有支持的文档（PDF及图片）"""
    # 确保目录存在
    Path(input_root).mkdir(parents=True, exist_ok=True)
    Path(output_root).mkdir(parents=True, exist_ok=True)
    
    # 遍历文件
    processed_files = 0
    failed_files = []
    
    all_files_to_process = []
    for ext in supported_extensions:
        all_files_to_process.extend(Path(input_root).glob(f"**/*{ext}"))

    total_files = len(all_files_to_process)

    if total_files > 0:
        for file_path in all_files_to_process:
            if process_document(str(file_path), output_root):
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

def cleanup_resources():
    """
    清理PaddleOCR相关资源，防止内存泄漏
    此函数应在程序结束前调用，以确保所有资源被正确释放
    """
    global _pipeline, _pipeline_initialized
    
    try:
        if _pipeline is not None:
            print("正在清理OCR资源...")
            
            # 尝试各种可能的清理方法
            cleanup_methods = ['cleanup', 'close', 'release', '__del__']
            for method_name in cleanup_methods:
                if hasattr(_pipeline, method_name):
                    try:
                        method = getattr(_pipeline, method_name)
                        if callable(method):
                            method()
                            print(f"  - 调用了 {method_name} 方法")
                            break
                    except Exception as e:
                        print(f"  - 调用 {method_name} 失败: {str(e)}")
                        continue
            
            # 清理可能的内部属性
            if hasattr(_pipeline, '__dict__'):
                for attr_name in list(_pipeline.__dict__.keys()):
                    try:
                        delattr(_pipeline, attr_name)
                    except:
                        pass
            
            # 设置为None
            _pipeline = None
            _pipeline_initialized = False
            
            # 强制垃圾回收
            gc.collect()
            
            print("OCR资源清理完成")
    except Exception as e:
        print(f"清理OCR资源时出错: {str(e)}")
    
    # 额外的清理步骤 - 尝试清理PaddlePaddle相关资源
    try:
        import paddle
        if hasattr(paddle, 'device') and hasattr(paddle.device, 'cuda'):
            if paddle.device.cuda.device_count() > 0:
                paddle.device.cuda.empty_cache()
                print("已清理CUDA缓存")
    except Exception as e:
        pass  # 忽略paddle相关的清理错误
        
if __name__ == "__main__":
    input_root = r"E:\Document\petrochina\知识问答工作流平台\doc_preparation\input"
    output_root = r"E:\Document\petrochina\知识问答工作流平台\doc_preparation\output"
    try:
        process_directory(input_root, output_root)
    finally:
        cleanup_resources()
