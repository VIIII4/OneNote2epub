import re
import os
import shutil
import sys
import win32com.client as win32 
import pywintypes 
import traceback
from xml.etree import ElementTree

OUTPUT_DIR = os.path.join(os.path.expanduser('~'), "Desktop", "OneNoteExport")
ASSETS_DIR = "assets"
PROCESS_RECYCLE_BIN = False
LOGFILE = 'onenote_to_docx.log'  # Set to None to disable logging
# For debugging purposes, set this variable to limit which pages are exported:
LIMIT_EXPORT = r''  # example: YourNotebook\Notes limits it to the Notes tab/page

def log(message):
    print(message)
    if LOGFILE is not None:
        with open(LOGFILE, 'a', encoding='UTF-8') as lf:
            lf.write(f'{message}\n')

def safe_str(name):
    return re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", name)

def should_handle(path):
    return path.startswith(LIMIT_EXPORT)

def handle_page(onenote, elem, path, i):
    safe_name = safe_str("%s_%s" % (str(i).zfill(3), elem.attrib['name']))
    if not should_handle(os.path.join(path, safe_name)):
        return

    full_path = os.path.join(OUTPUT_DIR, path)
    os.makedirs(full_path, exist_ok=True)
    safe_path = os.path.join(full_path, safe_name)
    path_docx = safe_path + '.docx'

    # Remove temp files if exist
    if os.path.exists(path_docx):
        os.remove(path_docx)

    try:
        # Create docx
        onenote.Publish(elem.attrib['ID'], path_docx, win32.constants.pfWord, "")
        log("Generating docx: %s" % path_docx)
    except pywintypes.com_error as e:
        log("!!WARNING!! Page Failed: %s" % path_docx)

def handle_element(onenote, elem, path='', i=0):
    if elem.tag.endswith('Notebook'):
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsChildren, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i)
    elif elem.tag.endswith('Section'):
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsPages, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i)
    elif elem.tag.endswith('SectionGroup') and (not elem.attrib['name'].startswith('OneNote_RecycleBin') or PROCESS_RECYCLE_BIN):
        hier2 = onenote.GetHierarchy(elem.attrib['ID'], win32.constants.hsSections, "")
        for i, c2 in enumerate(ElementTree.fromstring(hier2)):
            handle_element(onenote, c2, os.path.join(path, safe_str(elem.attrib['name'])), i)
    elif elem.tag.endswith('Page'):
        try:
            handle_page(onenote, elem, path, i)
        except:
            print("Page failed unexpectedly: %s" % path, file=sys.stderr)

if __name__ == "__main__":
    try:
        onenote = win32.gencache.EnsureDispatch("OneNote.Application.12")

        hier = onenote.GetHierarchy("", win32.constants.hsNotebooks, "")

        root = ElementTree.fromstring(hier)
        for child in root:
            handle_element(onenote, child)

    except pywintypes.com_error as e:
        traceback.print_exc()
        log("!!!Error!!! Hint: Make sure OneNote is open first.")