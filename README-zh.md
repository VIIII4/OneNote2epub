# OneNote2epub

这是一个将OneNote文件转换为epub格式的简易工具。

本项目的后半部分开发环境为Ubuntu系统，运行时只需授权即可。因此建议在Linux系统上运行后半部分流程，使用Windows会相对繁琐——需要关闭杀毒软件、以管理员权限运行程序，或在Windows安全中心进行相应授权（如果电脑未安装360安全软件）。

灵感源自[onenote-to-markdown](https://gitlab.com/pagekey/edu/onenote-to-markdown)

## 使用说明

1. 安装 Python 3.7+ （推荐3.8+版本）

2. 安装 [Pandoc](https://www.pandoc.org/installing.html) 并确保**已添加至系统环境变量PATH**

3. 安装依赖包
```bash
pip install -r requirements.txt
```

4. 确保**OneNote应用**正在运行

5. 运行onenote_to_docx.py
```bash
python onenote_to_docx.py
```
> 注意：有时需要直接在终端运行脚本（而非通过IDE），可尝试执行类似命令：`python 'D:\python\onenote2epub(+)\onenote_to_docx.py'`

6. 安装 [libreoffice](https://www.libreoffice.org/) 和 [Calibre](https://calibre-ebook.com/)

7. 修改coreConver.py中的**LIBREOFFICE_PATH**为你的libreoffice安装路径,windows的为其exe路径

8. 打开Calibre安装EpubMerge插件
![示例](https://raw.githubusercontent.com/VIIII4/OneNote2epub/master/images/20250416105139.png)

9. 运行main.py
```bash
python main.py
```

10. 在终端中输入位于**桌面**的**OneNoteExport**文件夹路径

即可在kindle上享受你的笔记啦！

## 已知缺陷

有点~~屎山~~

## 作者吐槽

其实非常不想用这种"缝合怪"方案——用Pandoc提取、LibreOffice转换、Calibre聚合。但单纯使用Pandoc时转换效果总是不尽人意，只能出此下策。
