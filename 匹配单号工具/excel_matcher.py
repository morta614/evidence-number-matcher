#!/usr/bin/env python3
"""
Excel 匹配工具
功能：
1. 读取总表 Excel
2. 读取 OCR 识别结果
3. 根据【单号+日期】匹配
4. 写入【取证来源】列
5. 导出结果

使用方法：
python excel_matcher.py --master 总表.xlsx --ocr complaint_records.xlsx --output 总表_已匹配.xlsx
"""

import pandas as pd
import argparse
from pathlib import Path
from tqdm import tqdm

def main():
    parser = argparse.ArgumentParser(description='匹配总表和 OCR 识别结果')
    parser.add_argument('--master', required=True, help='总表 Excel 文件')
    parser.add_argument('--ocr', required=True, help='OCR 识别结果 Excel')
    parser.add_argument('--output', default='总表_已匹配.xlsx', help='输出文件名')
    parser.add_argument('--id-col', default='投诉单号', help='总表中投诉单号的列名')
    parser.add_argument('--date-col', default='投诉日期', help='总表中投诉日期的列名')

    args = parser.parse_args()

    # 读取总表
    print(f"📂 读取总表：{args.master}")
    master_df = pd.read_excel(args.master)

    print(f"✅ 总表共 {len(master_df)} 行")

    # 检查列是否存在
    for col, name in [(args.id_col, '投诉单号'), (args.date_col, '投诉日期')]:
        if col not in master_df.columns:
            print(f"❌ 总表中没有列 '{col}'（{name}）")
            print(f"   可用列：{list(master_df.columns)}")
            return

    # 读取 OCR 结果
    print(f"📂 读取 OCR 结果：{args.ocr}")
    ocr_df = pd.read_excel(args.ocr)

    print(f"✅ OCR 结果共 {len(ocr_df)} 条")

    # 检查列
    if '单号' not in ocr_df.columns or '日期' not in ocr_df.columns:
        print(f"❌ OCR 结果中缺少 '单号' 或 '日期' 列")
        print(f"   可用列：{list(ocr_df.columns)}")
        return

    # 创建 OCR 查找字典（单号+日期 -> 来源）
    ocr_dict = {}
    for _, row in ocr_df.iterrows():
        key = (str(row['单号']), str(row['日期']))
        value = f"{row['来源压缩包']} - {row['截图名称']}"
        ocr_dict[key] = value

    # 匹配总表
    print(f"\n🔍 开始匹配...")
    matched_count = 0
    master_df['取证来源'] = '未匹配'

    for idx, row in tqdm(master_df.iterrows(), total=len(master_df), desc="匹配中"):
        key = (str(row[args.id_col]), str(row[args.date_col]))

        if key in ocr_dict:
            master_df.at[idx, '取证来源'] = ocr_dict[key]
            matched_count += 1

    # 导出结果
    output_path = Path(args.output)
    master_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\n✅ 匹配完成！")
    print(f"📊 总表记录数：{len(master_df)}")
    print(f"✅ 成功匹配：{matched_count} 条 ({matched_count/len(master_df)*100:.1f}%)")
    print(f"❌ 未匹配：{len(master_df) - matched_count} 条")
    print(f"📁 已导出到：{output_path.absolute()}")

    # 导出未匹配列表
    unmatched_df = master_df[master_df['取证来源'] == '未匹配']
    if not unmatched_df.empty:
        unmatched_path = Path('未匹配列表.xlsx')
        unmatched_df.to_excel(unmatched_path, index=False, engine='openpyxl')
        print(f"📋 未匹配列表已导出：{unmatched_path.absolute()}")

if __name__ == '__main__':
    main()
