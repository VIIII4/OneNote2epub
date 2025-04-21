from coreConver import Cmain
from merger import merge_epub_folder
import os,re,zipfile,tempfile
from pathlib import Path
import logging,shutil
from datetime import datetime
from tqdm import tqdm
def setup_logger():
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    logging.basicConfig(
        filename=log_dir/f"conversion_{timestamp}.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

def fix_unknown_titles(epub):
    """ Fix 'Unknown Title' entries in EPUB toc.ncx files by extracting proper titles from content files
    Uses regex instead of XML libraries

    Args:
        epub (str): Path to the EPUB file
    Returns:
        int: Number of titles fixed
    """
    temp_dir = tempfile.mkdtemp()
    logging.info(f"Created temporary directory: {temp_dir}")
    with zipfile.ZipFile(epub, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
        logging.info(f"Extracted EPUB to {temp_dir}")
    epub_root = temp_dir
    toc_path = os.path.join(epub_root, 'toc.ncx')
    try:
        if not os.path.exists(toc_path):
            logging.warning(f"toc.ncx not found in {epub}")
            return 0

        # Read the toc.ncx file
        with open(toc_path, 'r', encoding='utf-8') as file:
            toc_content = file.read()
        logging.info(f"Read toc.ncx content from {toc_path}")

        # Pattern to find navPoint blocks
        nav_point_pattern = re.compile(r'<navPoint\s+[^>]*id="([^"]+)"[^>]*>.*?</navPoint>', re.DOTALL)
        nav_points = nav_point_pattern.findall(toc_content)

        fixed_count = 0

        # For each navPoint, check if it contains "Unknown Title"
        for nav_point_id in nav_points:
            # Extract the complete navPoint block
            nav_pattern = re.compile(r'(<navPoint\s+[^>]*id="' + re.escape(nav_point_id) +
                                     r'"[^>]*>.*?</navPoint>)', re.DOTALL)
            nav_match = nav_pattern.search(toc_content)

            if nav_match:
                nav_block = nav_match.group(1)

                # Check if this block contains "Unknown Title"
                text_pattern = re.compile(r'<text>Unknown Title</text>')
                if text_pattern.search(nav_block):
                    # Extract the content src attribute
                    content_pattern = re.compile(r'<content\s+src="([^"]+)"')
                    content_match = content_pattern.search(nav_block)

                    if content_match:
                        content_src = content_match.group(1)
                        content_path = os.path.join(epub_root, content_src)

                        try:
                            # Read the content file
                            with open(content_path, 'r', encoding='utf-8') as file:
                                content_data = file.read()
                            logging.info(f"Read content from {content_path}")

                            # Extract the title
                            title_pattern = re.compile(r'<title>([^<]+)</title>')
                            title_match = title_pattern.search(content_data)

                            if title_match:
                                actual_title = title_match.group(1)

                                # Replace "Unknown Title" with the actual title
                                updated_nav_block = re.sub(
                                    r'<text>Unknown Title</text>',
                                    f'<text>{actual_title}</text>',
                                    nav_block
                                )

                                # Update the full content
                                toc_content = toc_content.replace(nav_block, updated_nav_block)
                                fixed_count += 1
                                logging.info(f"Fixed: \"{actual_title}\" ({content_src})")
                            else:
                                logging.warning(f"No title found in {content_path}")
                        except Exception as e:
                            logging.error(f"Error processing content file {content_path}: {str(e)}")

        # Write the updated toc.ncx file if changes were made
        if fixed_count > 0:
            with open(toc_path, 'w', encoding='utf-8') as file:
                file.write(toc_content)
            logging.info(f"Updated toc.ncx with {fixed_count} fixed titles")
        else:
            logging.info('No "Unknown Title" entries found in toc.ncx')

        # 重新打包为 EPUB 文件
        file_path = epub
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)  # 获取相对路径
                    zip_ref.write(file_path, arcname)
        logging.info(f"Repackaged EPUB file: {file_path}")

        return fixed_count

    except Exception as e:
        logging.error(f"Error processing toc.ncx: {str(e)}")
        raise
    finally:
        # 删除临时文件夹
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logging.info(f"Deleted temporary directory: {temp_dir}")

def delete_folder_contents(folder_path: str):
    """
    安全删除指定文件夹内的所有内容（保留文件夹本身）
    
    :param folder_path: 要清空的文件夹路径
    """
    target_dir = Path(folder_path).resolve()
    
    if not target_dir.exists():
        logging.warning(f"目标文件夹不存在: {target_dir}")
        return

    if not target_dir.is_dir():
        logging.error(f"路径不是文件夹: {target_dir}")
        raise ValueError("必须提供文件夹路径")

    logging.info(f"开始清理文件夹: {target_dir}")
    
    try:
        # 删除所有子文件和子目录
        for item in tqdm(list(target_dir.glob("*")), desc="删除文件中", unit="item"):
            try:
                if item.is_file():
                    item.unlink()
                    logging.debug(f"已删除文件: {item}")
                elif item.is_dir():
                    shutil.rmtree(item)
                    logging.debug(f"已删除目录: {item}")
            except Exception as e:
                logging.error(f"删除失败 {item}: {str(e)}")
                raise

        logging.info(f"文件夹清理完成: {target_dir}")
    except Exception as e:
        logging.error(f"清理过程中断: {str(e)}")
        raise
def find_docx_folders(root_dir='.'):
    """
    查找包含docx文件的文件夹路径
    
    :param root_dir: 项目根目录路径（默认为当前目录）
    :return: 包含docx文件的文件夹路径列表
    """
    docx_folders = []
    root_path = Path(root_dir).resolve()

    for root, _, files in os.walk(root_path):
        # 检查当前目录是否包含docx文件
        if any(f.lower().endswith('.docx') for f in files):
            # 将绝对路径转换为字符串并去重
            folder_path = str(Path(root).resolve())
            if folder_path not in docx_folders:
                docx_folders.append(folder_path)

    return docx_folders


def ConvertFirst(docx_folders):
    EpubList = []
    for folder in tqdm(docx_folders, desc="Converting folders"):
        output_dir = os.path.join('','internEpubs',f'{Path(folder).name}')
        logging.info(f"Converting {folder}...")
        Cmain(source_folder=folder,output_folder=output_dir) 
        EpubList.append(output_dir)
    return EpubList

def MergeEpub(EpubList):
    for folder in tqdm(EpubList, desc="Merging EPUBs"):
        logging.info(f"Merging {folder}...")
        output = os.path.join('','finalEpubs',f'{Path(folder).name}.epub')
        merge_epub_folder(folder,
                         output=output,
                         title=f'{Path(folder).name}',
                         author='VIIII4258',
                         sort="date_reverse"
                         )
    delete_folder_contents(os.path.join('','internEpubs'))


if __name__ == "__main__":
    os.makedirs(os.path.join('','internEpubs'), exist_ok=True)
    os.makedirs(os.path.join('','finalEpubs'), exist_ok=True)
    setup_logger()
    delete_folder_contents(os.path.join('','finalEpubs'))
    delete_folder_contents(os.path.join('','internEpubs'))
    try:
        logging.info("Program started")
        root_dir = str(input("请输入项目根目录路径："))
        
        logging.info("Searching for DOCX folders...")
        docx_folders = find_docx_folders(root_dir)
        logging.info(f"Found {len(docx_folders)} folders with DOCX files")
        
        logging.info("Starting conversion...")
        EpubList = ConvertFirst(docx_folders)
        
        logging.info("Merging EPUB files...")
        MergeEpub(EpubList)
        
        logging.info("Process completed successfully,see finalEpubs for result")
        key = input("想要继续将这些书（onenote所有笔记)合成一本吗？(~~有概率死机~~）(y/n)")
        if key == 'y' or key == 'Y':
            shuming = input("请输入书名")
            zuozhe = input("请输入作者")
            merge_epub_folder(os.path.join('','finalEpubs'),
                         output=os.path.join('',f'{shuming}.epub'),
                         title=str(shuming),
                         author=str(zuozhe),
                         sort="date_reverse"
                         )
            fix_unknown_titles(os.path.join('',f'{shuming}.epub'))
            delete_folder_contents(os.path.join('','finalEpubs'))
            logging.info("Process completed successfully,see this root for result")
    except Exception as e:
        logging.error(f"Error occurred: {str(e)}", exc_info=True)
        raise
    



    