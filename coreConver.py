import os,re,zipfile,tempfile
import subprocess
import logging,shutil
from typing import List

# 默认配置
DEFAULT_CONFIG = {
    "SOURCE_FOLDER": "/home/viiii4258/桌面/OneNoteExport/Motion/日志",
    "DELETE_ORIGINAL": False,
    "LIBREOFFICE_PATH": "/usr/bin/libreoffice",
    "OUTPUT_FOLDER": "/home/viiii4258/onenote2epub(+)/internEpubs"
}

# 初始化日志（保持全局）
logging.basicConfig(
    filename="docx_conversion.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# fixTitle

def insert_title_into_head(file_path,title=None):
    """
    提取指定标签中的内容并插入到<head>中<link>元素后。
    如果<head>中已有<title>标签，则覆盖其内容。
    
    :param file_path: XHTML 文件路径
    """
    # 正则表达式匹配目标内容
    target_pattern = re.compile(r'<p class="para0">(.*?)</p>', re.DOTALL)
    head_pattern = re.compile(r'(<head>.*?</head>)', re.DOTALL)
    link_pattern = re.compile(r'(<link[^>]*>)')
    title_pattern = re.compile(r'<title>(.*?)</title>')  # 匹配已有<title>标签


    try:
        # 分块读取文件内容，避免一次性加载大文件
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # 查找目标内容
        if title is None:
            match = target_pattern.search(content)
            if not match:
                print("未找到符合条件的目标内容。")
                return

            # 提取匹配的内容并去除标签
            title_content = re.sub(r'<[^>]+>', '', match.group(1))
            title_tag = f"<title>{title_content}</title>"
        else:
            title_tag = f"<title>{title}</title>"
        
        # 查找 <head> 部分
        head_match = head_pattern.search(content)
        if not head_match:
            print("未找到 <head> 标签。")
            return
        
        head_content = head_match.group(1)

        # 检查是否已有<title>标签
        existing_title_match = title_pattern.search(head_content)
        if existing_title_match:
            # 如果已有<title>标签，替换其内容
            old_title_tag = existing_title_match.group(0)
            new_head_content = head_content.replace(old_title_tag, title_tag)
        else:
            # 如果没有<title>标签，插入到<link>元素后面
            link_match = link_pattern.search(head_content)
            if not link_match:
                print("未找到 <link> 元素。")
                return

            link_tag = link_match.group(1)
            new_head_content = head_content.replace(link_tag, f"{link_tag}\n{title_tag}")

        # 替换<head>部分
        new_content = content.replace(head_content, new_head_content)

        # 写回文件
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(new_content)

        print(f"成功插入或更新标题: {title_tag}")

    except Exception as e:
        print(f"处理文件时发生错误: {e}")

def anlyze_epub(file_path):
    '''
    输入epub文件路径，解压epub文件，到一个临时文件夹中，对临时文件夹进行遍历以对每个xhtml文件进行修改
    修改后将临时文件夹重新打包成epub文件并覆盖原文件
    最后删除临时文件夹
    '''
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 解压 EPUB 文件到临时文件夹
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # 遍历临时文件夹，找到所有 .xhtml 文件并调用 insert_title_into_head 进行修改
        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.endswith('.xhtml') and file != 'toc.xhtml':
                    xhtml_file_path = os.path.join(root, file)
                    insert_title_into_head(xhtml_file_path)
        
        # 重新打包为 EPUB 文件
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)  # 获取相对路径
                    zip_ref.write(file_path, arcname)
        
        print(f"成功更新并重新打包 EPUB 文件: {file_path}")
    
    except Exception as e:
        print(f"处理 EPUB 文件时发生错误: {e}")
    
    finally:
        # 删除临时文件夹
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)


# /

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

    for epub in os.listdir(output_folder):
        if epub.endswith(".epub"):
            anlyze_epub(os.path.join(output_folder, epub))

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