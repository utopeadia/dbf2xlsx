#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DBF to Excel Converter

This script converts DBF files to Excel files.

Author: Wu Shuhui
"""

import pandas as pd
from dbfread import DBF, DBFNotFound, MissingMemoFile
import argparse
import os
import sys
from tqdm import tqdm
import logging

# 配置日志记录
logging.basicConfig(
    filename="dbf_to_excel.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

SOFTWARE_NAME = "DBF2XLSX"
VERSION = "0.1"
AUTHOR = "Wu Shuhui"


def dbf_to_excel(dbf_file_path, excel_file_path, encoding):
    """
    将 DBF 文件转换为 Excel 文件。

    Args:
      dbf_file_path: DBF 文件的路径。
      excel_file_path: 输出 Excel 文件的路径。
      encoding: 文件编码，可选。
    """
    try:
        if encoding is None:
            encodings_to_try = ["gbk", "utf-8", "gb2312", "gb18030", "latin1"]
            for enc in encodings_to_try:
                try:
                    dbf = DBF(dbf_file_path, encoding=enc, load=True)
                    df = pd.DataFrame(dbf.records)
                    df.to_excel(excel_file_path, index=False)
                    logging.info(
                        f"成功使用编码 '{enc}' 将 {dbf_file_path} 转换为 {excel_file_path}"
                    )
                    return
                except UnicodeDecodeError:
                    logging.warning(
                        f"尝试使用编码 '{enc}' 解码文件 '{dbf_file_path}' 失败"
                    )
                except Exception as e:
                    logging.error(
                        f"使用编码 '{enc}' 转换文件 '{dbf_file_path}' 失败: {e}"
                    )

            logging.error(f"错误：无法找到合适的编码来解码文件 '{dbf_file_path}'")

        else:
            dbf = DBF(dbf_file_path, encoding=encoding, load=True)
            df = pd.DataFrame(dbf.records)
            df.to_excel(excel_file_path, index=False)
            logging.info(f"成功将 {dbf_file_path} 转换为 {excel_file_path}")

    except DBFNotFound:
        logging.error(f"错误：找不到文件 '{dbf_file_path}'")
    except MissingMemoFile:
        logging.error(f"错误：缺少 Memo 文件 '{dbf_file_path}'")
    except Exception as e:
        logging.error(f"转换文件 '{dbf_file_path}' 失败: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="将指定 DBF 文件或目录下的 DBF 文件转换为 Excel 文件。"
    )

    parser.add_argument("input_path", help="DBF 文件或包含 DBF 文件的目录")
    parser.add_argument(
        "-e",
        "--encoding",
        help="DBF 文件的编码 (例如 gbk, utf-8, latin1 等，不指定则自动检测)",
        default=None,
    )
    parser.add_argument(
        "-o",
        "--output_dir",
        help="输出 Excel 文件的目录，默认为输入路径所在目录",
        default=None,
    )

    args = parser.parse_args()

    # 获取输入路径和输出目录
    input_path = args.input_path
    output_dir = args.output_dir

    if not os.path.exists(input_path):
        print(f"错误：输入路径 '{input_path}' 不存在。")
        sys.exit(1)

    if os.path.isfile(input_path) and input_path.lower().endswith(".dbf"):
        # 如果是单个 DBF 文件
        dbf_files = [input_path]
        if output_dir is None:
            output_dir = os.path.dirname(input_path)
    elif os.path.isdir(input_path):
        # 如果是目录
        dbf_files = [
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".dbf")
        ]
        if output_dir is None:
            output_dir = input_path
    else:
        print(f"错误：输入路径 '{input_path}' 不是有效的 DBF 文件或目录。")
        sys.exit(1)

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 显示软件信息
    print(f"== {SOFTWARE_NAME} v{VERSION} ==")
    print(f"== 作者：{AUTHOR} ==")
    print("*" * 30)

    for dbf_file_path in tqdm(dbf_files, desc="转换进度"):
        excel_file_name = os.path.splitext(os.path.basename(dbf_file_path))[0] + ".xlsx"
        excel_file_path = os.path.join(output_dir, excel_file_name)

        # 执行转换
        dbf_to_excel(dbf_file_path, excel_file_path, args.encoding)


if __name__ == "__main__":
    main()
