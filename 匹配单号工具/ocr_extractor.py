#!/usr/bin/env python3
"""
OCR 识别工具（增强版）
功能：
1. 读取导出的图片列表 CSV
2. 批量 OCR 识别每张图片
3. 提取投诉单号和日期
4. 同时提取截图名称、反馈产品、作品名称
5. 导出为 Excel 表格
6. 记录来源压缩包

使用方法：
python ocr_extractor.py --input selected_images.csv --output complaint_records.xlsx
"""

import pytesseract
from PIL import Image
import pandas as pd
import re
from pathlib import Path
import argparse
from tqdm import tqdm

def extract_complaint_info(image_path):
    """
    从截图中提取投诉信息
    返回：{
        '截图名称': 文件名,
        '反馈产品': 产品名,
        '作品名称': 作品名称,
        'records': [(单号, 日期), ...]
    }
    """
    try:
        # 读取图片
        img = Image.open(image_path)

        # OCR 识别
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')

        # 提取截图名称
        截图名称 = image_path.name

        # 提取反馈产品（通常在"反馈产品"后面）
        反馈产品 = "未识别"
        product_match = re.search(r'反馈产品[^\n]*\n([^\n]+)', text)
        if product_match:
            product_text = product_match.group(1).strip()
            # 清理多余字符
            product_text = re.sub(r'[^\w\u4e00-\u9fff]', '', product_text)
            if product_text:
                反馈产品 = product_text

        # 提取作品名称（通常在"作品名称"后面）
        作品名称 = "未识别"
        work_match = re.search(r'作品名称[^\n]*\n([^\n]+)', text)
        if work_match:
            work_text = work_match.group(1).strip()
            # 清理多余字符
            work_text = re.sub(r'[^\w\u4e00-\u9fff]', '', work_text)
            if work_text:
                作品名称 = work_text

        # 提取所有 8 位数字（单号）
        numbers = re.findall(r'\d{8}', text)

        # 提取日期（格式：20xx-xx-xx）
        dates = re.findall(r'20\d{2}-\d{2}-\d{2}', text)

        # 匹配单号和日期（假设按表格顺序一一对应）
        records = []
        if numbers and dates:
            min_len = min(len(numbers), len(dates))
            for i in range(min_len):
                records.append({
                    '单号': numbers[i],
                    '日期': dates[i]
                })

        return {
            '截图名称': 截图名称,
            '反馈产品': 反馈产品,
            '作品名称': 作品名称,
            'records': records
        }

    except Exception as e:
        print(f"❌ 识别失败 {image_path}: {e}")
        return {
            '截图名称': image_path.name,
            '反馈产品': '识别失败',
            '作品名称': '识别失败',
            'records': []
        }

def main():
    parser = argparse.ArgumentParser(description='OCR 批量识别投诉截图')
    parser.add_argument('--input', required=True, help='输入的 CSV 文件（导出的图片列表）')
    parser.add_argument('--output', default='complaint_records.xlsx', help='输出的 Excel 文件')
    parser.add_argument('--source-col', default='完整路径', help='CSV 中图片路径的列名')

    args = parser.parse_args()

    # 读取 CSV
    print(f"📂 读取文件：{args.input}")
    df = pd.read_csv(args.input)

    if args.source_col not in df.columns:
        print(f"❌ 列名 '{args.source_col}' 不存在，可用列：{list(df.columns)}")
        return

    print(f"✅ 共 {len(df)} 张图片待识别")

    # 批量识别
    all_records = []

    for idx, row in tqdm(df.iterrows(), total=len(df), desc="OCR 识别中"):
        image_path = Path(row[args.source_col])

        if not image_path.exists():
            print(f"⚠️ 文件不存在：{image_path}")
            continue

        # 识别图片
        info = extract_complaint_info(image_path)

        # 为每条记录添加元数据
        for record in info['records']:
            record['截图名称'] = info['截图名称']
            record['反馈产品'] = info['反馈产品']
            record['作品名称'] = info['作品名称']
            record['来源压缩包'] = row.get('来源压缩包', '未知')

        all_records.extend(info['records'])

    # 转换为 DataFrame
    result_df = pd.DataFrame(all_records)

    if result_df.empty:
        print("❌ 未识别到任何投诉记录")
        return

    # 调整列顺序（把新增的三列放在最前面）
    columns_order = ['截图名称', '反馈产品', '作品名称', '单号', '日期', '来源压缩包']
    result_df = result_df[columns_order]

    # 去重（同一单号+日期可能出现在多张截图）
    result_df = result_df.drop_duplicates(subset=['单号', '日期'])

    # 导出 Excel
    output_path = Path(args.output)
    result_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\n✅ 识别完成！")
    print(f"📊 共识别 {len(result_df)} 条投诉记录")
    print(f"📁 已导出到：{output_path.absolute()}")

    # 统计信息
    print(f"\n📈 统计：")
    print(f"  - 原始截图数：{len(df)} 张")
    print(f"  - 识别出记录数：{len(all_records)} 条")
    print(f"  - 去重后记录数：{len(result_df)} 条")
    print(f"  - 涉及压缩包：{result_df['来源压缩包'].nunique()} 个")

    # 显示前几条记录预览
    print(f"\n📋 数据预览（前5条）：")
    print(result_df.head().to_string(index=False))

if __name__ == '__main__':
    main()
