#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calibre EpubMerge自动化脚本
用于合并指定文件夹内的所有epub文件
"""

import os
import subprocess
import argparse
import glob
import shutil
from pathlib import Path


def find_epub_files(folder_path):
    """查找指定文件夹中的所有epub文件"""
    epub_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith('.epub'):
                epub_files.append(os.path.join(root, file))
    return sorted(epub_files)  # 按文件名排序


def merge_epub_files(epub_files, output_file, calibre_path="calibre-debug", title=None, author=None, 
                     description=None, tags=None, cover_img=None, titles_nav_points=None, 
                     nav_points_insert=None, source_nav_rule=None):
    """使用Calibre的EpubMerge插件合并epub文件"""
    if not epub_files:
        print("没有找到epub文件")
        return False
    
    print(f"找到 {len(epub_files)} 个epub文件:")
    for i, epub in enumerate(epub_files, 1):
        print(f"{i}. {os.path.basename(epub)}")
    
    # 构建命令行调用
    cmd = [
        calibre_path,
        "--run-plugin",
        "EpubMerge",
        "--"
    ]
    
    # 添加各种可选参数
    if output_file:
        cmd.extend(["--output", output_file])
    
    if title:
        cmd.extend(["--title", title])
    
    if author:
        cmd.extend(["--author", author])
    
    if description:
        cmd.extend(["--description", description])
    
    if tags:
        cmd.extend(["--tags", tags])
    
    if cover_img:
        cmd.extend(["--coverimg", cover_img])
    
    if titles_nav_points is not None:
        cmd.extend(["--titlesnavpoints", str(titles_nav_points)])
    
    if nav_points_insert is not None:
        cmd.extend(["--navpointsinsert", str(nav_points_insert)])
    
    if source_nav_rule is not None:
        cmd.extend(["--sourcenavrule", str(source_nav_rule)])
    
    # 添加所有输入epub文件
    cmd.extend(epub_files)
    
    try:
        print("\n执行合并命令...")
        print(' '.join(cmd))
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("合并成功!")
        print(f"输出文件: {output_file}")
        return True
    
    except subprocess.CalledProcessError as e:
        print(f"合并失败: {e}")
        if e.stderr:
            print(f"错误输出: {e.stderr}")
        return False




def merge_epub_folder(
    folder,
    output=None,
    title=None,
    author=None,
    description=None,
    tags=None,
    cover_img=None,
    titles_nav_points=None,
    nav_points_insert=None,
    source_nav_rule=None,
    calibre_path="calibre-debug",
    sort="name"
):
    
    
    """函数式调用入口"""
    # 确保文件夹存在
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"文件夹 '{folder}' 不存在")

    # 查找epub文件
    epub_files = find_epub_files(folder)

    if not epub_files:
        raise FileNotFoundError(f"在 '{folder}' 中没有找到epub文件")

    # 文件排序逻辑
    sort_options = {
        "name": lambda x: os.path.basename(x).lower(),
        "name_reverse": lambda x: os.path.basename(x).lower(),
        "size": lambda x: os.path.getsize(x),
        "size_reverse": lambda x: os.path.getsize(x),
        "date": lambda x: os.path.getmtime(x),
        "date_reverse": lambda x: os.path.getmtime(x)
    }
    reverse = sort.endswith("_reverse")
    epub_files.sort(
        key=sort_options[sort],
        reverse=reverse
    )

    # 处理默认输出路径
    if not output:
        folder_name = os.path.basename(os.path.abspath(folder))
        output = f"{folder_name}_merged.epub"

    # 路径处理
    output = os.path.abspath(output)
    output_dir = os.path.dirname(output)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 调用核心合并函数
    return merge_epub_files(
        epub_files=epub_files,
        output_file=output,
        calibre_path=calibre_path,
        title=title,
        author=author,
        description=description,
        tags=tags,
        cover_img=cover_img,
        titles_nav_points=titles_nav_points,
        nav_points_insert=nav_points_insert,
        source_nav_rule=source_nav_rule
    )





def main():
    parser = argparse.ArgumentParser(description="使用Calibre的EpubMerge插件合并文件夹中的epub文件")
    parser.add_argument("folder", help="包含epub文件的文件夹路径")
    parser.add_argument("-o", "--output", help="输出的epub文件路径", default=None)
    parser.add_argument("-t", "--title", help="合并后的epub标题", default=None)
    parser.add_argument("-a", "--author", help="合并后的epub作者", default=None)
    parser.add_argument("-d", "--description", help="合并后的epub描述", default=None)
    parser.add_argument("-g", "--tags", help="合并后的epub标签", default=None)
    parser.add_argument("-i", "--cover-img", help="封面图片路径", default=None)
    parser.add_argument("-m", "--titles-nav-points", type=int, choices=[0, 1], 
                        help="为标题插入导航点 (0=否, 1=是)", default=None)
    parser.add_argument("-n", "--nav-points-insert", type=int, choices=[0, 1], 
                        help="在源书籍开始处插入导航点 (0=否, 1=是)", default=None)
    parser.add_argument("-s", "--source-nav-rule", type=int, choices=[0, 1, 2], 
                        help="如何处理目录中的书籍链接 (0=保持原样, 1=修改为合并后的文件, 2=删除)", default=None)
    parser.add_argument("--calibre-path", help="calibre-debug的路径", default="calibre-debug")
    parser.add_argument("--sort", choices=["name", "name_reverse", "size", "size_reverse", "date", "date_reverse"],
                        help="文件排序方式", default="name")
    args = parser.parse_args()
    
    # 确保文件夹存在
    if not os.path.isdir(args.folder):
        print(f"错误: 文件夹 '{args.folder}' 不存在")
        return 1
    
    # 查找epub文件
    epub_files = find_epub_files(args.folder)
    
    if not epub_files:
        print(f"在 '{args.folder}' 中没有找到epub文件")
        return 1
    
    # 根据指定的方式排序文件
    if args.sort == "name":
        epub_files.sort(key=lambda x: os.path.basename(x).lower())
    elif args.sort == "name_reverse":
        epub_files.sort(key=lambda x: os.path.basename(x).lower(), reverse=True)
    elif args.sort == "size":
        epub_files.sort(key=lambda x: os.path.getsize(x))
    elif args.sort == "size_reverse":
        epub_files.sort(key=lambda x: os.path.getsize(x), reverse=True)
    elif args.sort == "date":
        epub_files.sort(key=lambda x: os.path.getmtime(x))
    elif args.sort == "date_reverse":
        epub_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    # 如果没有指定输出文件名，则使用文件夹名作为基础
    if not args.output:
        folder_name = os.path.basename(os.path.abspath(args.folder))
        args.output = f"{folder_name}_merged.epub"
    
    # 确保输出路径是绝对路径
    if not os.path.isabs(args.output):
        args.output = os.path.abspath(args.output)
    
    # 确保输出目录存在
    output_dir = os.path.dirname(args.output)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    args = parser.parse_args()
    
    try:
        success = merge_epub_folder(
            folder=args.folder,
            output=args.output,
            title=args.title,
            author=args.author,
            description=args.description,
            tags=args.tags,
            cover_img=args.cover_img,
            titles_nav_points=args.titles_nav_points,
            nav_points_insert=args.nav_points_insert,
            source_nav_rule=args.source_nav_rule,
            calibre_path=args.calibre_path,
            sort=args.sort
        )
        return 0 if success else 1
    except Exception as e:
        print(str(e))
        return 1


if __name__ == "__main__":
    exit(main())


'''
from epub_merge import merge_epub_folder

success = merge_epub_folder(
    folder="/path/to/books",
    output="merged.epub",
    title="合集",
    author="多人",
    sort="date_reverse"
)
'''