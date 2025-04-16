# OneNote2epub

This is a simple tool to convert OneNote files to epub format.

The latter part of this project was developed on Ubuntu and you just need to authorize the program when running it. Therefore, it is recommended to run the latter part on a Linux system. Using Windows would be more cumbersome. You need to turn off the antivirus software, run the program with administrator privileges, or make corresponding authorizations in the Windows Security Center (if you don't have 360 security software on your computer).

It is inspired by[onenote-to-markdown](https://gitlab.com/pagekey/edu/onenote-to-markdown)

## Usage

1. Install Python 3.7+  (Python 3.8+ is recommended)

2. Install [Pandoc](https://www.pandoc.org/installing.html) and **add it to PATH**

3. Install the required packages

```bash
pip install -r requirements.txt
```

4. make sure your **OneNote** application is running

5. Run the onenote_to_docx.py

```bash
python onenote_to_docx.py
```

>sometimes you need to run the onenote_to_docx.py **aside** of the IDE to avoid the error.Something like run this command directly in the terminal: `python 'D:\python\onenote2epub(+)\onenote_to_docx.py'`

6. Install [libreoffice](https://www.libreoffice.org/) and [Calibre](https://calibre-ebook.com/)

7. change the **LIBREOFFICE_PATH** in coreConver.py to your libreoffice path

8. Open Calibre and install the EpubMerge plugin.
![图片描述](https://raw.githubusercontent.com/用户名/仓库名/master/images/图片名.png)

9. Run the main.py

```bash
python main.py
```

10. put the path of **OneNoteExport** where it will exit on your **Desktop** in the terminal

enjoy your notes on kindle!

## Defect

The high - level part of the navigation directory in the epub file is normal, but the names of specific chapters are currently all empty (this will be fixed later when there is time). However, the content display is fine.

## Roast

I really don't want to use such a "stitched" approach of extracting with Pandoc, converting with LibreOffice, and aggregating with Calibre. But the results are always unsatisfactory when using only Pandoc.