import os
import subprocess
import logging
from typing import List

# 默认配置
DEFAULT_CONFIG = {
    "SOURCE_FOLDER": r"D:\python\onenote2epub(+)\OneNoteExport\调查报告\这是一个令人疑惑的星球",
    "DELETE_ORIGINAL": False,
    "LIBREOFFICE_PATH": r"C:\Program Files\LibreOffice\program\soffice.exe",
    "OUTPUT_FOLDER": r"D:\python\onenote2epub(+)\internEpubs"
}

# 初始化日志（保持全局）
logging.basicConfig(
    filename="docx_conversion.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def get_docx_files(folder: str) -> List[str]:
    """获取指定文件夹下所有 docx 文件（排除临时文件）"""
    return [
        os.path.join(folder, f) for f in os.listdir(folder)
        if f.endswith(".docx") and not f.startswith("~")
    ]

def convert_docx_to_epub(
    file_path: str,
    output_folder: str = DEFAULT_CONFIG["OUTPUT_FOLDER"],
    delete_original: bool = DEFAULT_CONFIG["DELETE_ORIGINAL"],
    libreoffice_path: str = DEFAULT_CONFIG["LIBREOFFICE_PATH"]
) -> bool:
    """使用 LibreOffice 转换单个 docx 文件到 epub"""
    try:
        os.makedirs(output_folder, exist_ok=True)
        
        # 生成输出路径
        file_name = os.path.basename(file_path)
        epub_name = os.path.splitext(file_name)[0] + ".epub"
        epub_path = os.path.join(output_folder, epub_name)
        
        command = [
            libreoffice_path,
            "--headless",
            "--convert-to", "epub",
            "--outdir", output_folder,
            file_path
        ]
        
        result = subprocess.run(
            command,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        if result.returncode == 0:
            actual_epub = os.path.join(output_folder, epub_name)
            if os.path.exists(actual_epub):
                logging.info(f"成功转换: {file_path} -> {actual_epub}")
                if delete_original:
                    os.remove(file_path)
                    logging.info(f"已删除原文件: {file_path}")
                return True
            
            logging.warning(f"转换成功但文件未生成: {actual_epub}")
            return False
        return False

    except subprocess.CalledProcessError as e:
        logging.error(f"转换失败: {file_path} - {e.stderr}")
    except Exception as e:
        logging.error(f"处理文件出错: {file_path} - {str(e)}")
    return False

def Cmain(
    source_folder: str = DEFAULT_CONFIG["SOURCE_FOLDER"],
    output_folder: str = DEFAULT_CONFIG["OUTPUT_FOLDER"],
    delete_original: bool = DEFAULT_CONFIG["DELETE_ORIGINAL"],
    libreoffice_path: str = DEFAULT_CONFIG["LIBREOFFICE_PATH"]
):
    if not os.path.exists(source_folder):
        logging.error(f"源文件夹不存在: {source_folder}")
        return
    
    docx_files = get_docx_files(source_folder)
    if not docx_files:
        logging.info("未发现可转换的 docx 文件")
        return
    
    try:
        os.makedirs(output_folder, exist_ok=True)
    except Exception as e:
        logging.error(f"无法创建输出目录: {output_folder} - {str(e)}")
        return
    
    print(f"开始处理 {len(docx_files)} 个 docx 文件...")
    success_count = 0
    
    for file in docx_files:
        if convert_docx_to_epub(
            file,
            output_folder=output_folder,
            delete_original=delete_original,
            libreoffice_path=libreoffice_path
        ):
            success_count += 1
    
    print(f"\n处理完成！成功转换 {success_count}/{len(docx_files)} 个文件")
    print(f"详细日志见: {os.path.abspath('docx_conversion.log')}")

if __name__ == "__main__":
    # 使用默认配置运行
    Cmain()





'''
# 使用自定义配置
main(
    source_folder="/custom/source",
    output_folder="/custom/output",
    delete_original=True,
    libreoffice_path="/custom/path/libreoffice"
)

# 仅修改部分配置
main(delete_original=True)
'''