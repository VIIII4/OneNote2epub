from coreConver import Cmain
from merger import merge_epub_folder
import os
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
        os.makedirs(output_dir, exist_ok=True)
        logging.info(f"Converting {folder}...")
        Cmain(source_folder=folder,output_folder=output_dir) 
        EpubList.append(output_dir)

    return EpubList

def MergeEpub(EpubList):
    output_dir = os.path.join('','finalEpubs')
    os.makedirs(output_dir, exist_ok=True)
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
        key = input("想要继续将这些书（onenote所有笔记)合成一本吗？(~~如果你电脑太烂那有概率死机~~）(y/n)")
        if key == 'y' or key == 'Y':
            shuming = input("请输入书名（可随意）")
            zuozhe = input("请输入作者（可随意）")
            merge_epub_folder(os.path.join('','finalEpubs'),
                         output=os.path.join('',f'{shuming}.epub'),
                         title=str(shuming),
                         author=str(zuozhe),
                         sort="date_reverse"
                         )
            delete_folder_contents(os.path.join('','finalEpubs'))
            logging.info("Process completed successfully,see this root for result")
    except Exception as e:
        logging.error(f"Error occurred: {str(e)}", exc_info=True)
        raise
    



    