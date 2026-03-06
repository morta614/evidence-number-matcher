#!/usr/bin/env python3
"""
取证单号匹配工具 - 便携版
自动使用内置的 Tesseract OCR，无需安装
"""

import streamlit as st
import zipfile
import os
from pathlib import Path
import pandas as pd
from PIL import Image
import base64
from io import BytesIO
import pytesseract
import re
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import sys

# ============ 便携版配置 ============
# 自动检测并使用项目内的 Tesseract
script_dir = Path(__file__).parent.absolute()

# 设置 Tesseract 路径（便携版）
tesseract_path = script_dir / "tesseract" / "tesseract.exe"

if tesseract_path.exists():
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)
    # 设置 tessdata 环境变量
    os.environ['TESSDATA_PREFIX'] = str(script_dir / "tesseract" / "tessdata")
else:
    st.error(f"❌ 找不到 Tesseract OCR！请确保 tesseract 文件夹在 {script_dir}")
    st.stop()

# 设置页面
st.set_page_config(
    page_title="取证单号匹配工具",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 侧边栏
st.sidebar.title("📋 取证单号匹配工具")
st.sidebar.markdown("""
### 功能介绍
- **取证预览**：上传压缩包，预览并勾选需要识别的截图
- **OCR 识别**：批量识别截图中的单号和日期
- **Excel 匹配**：匹配总表和OCR结果，标记取证来源
""")

# 创建临时目录
TEMP_DIR = script_dir / "temp_evidence"
TEMP_DIR.mkdir(exist_ok=True)

# 创建标签页
tab1, tab2, tab3 = st.tabs(["📸 取证预览", "🔍 OCR 识别", "📊 Excel 匹配"])

#============ Tab 1: 取证预览 ============
with tab1:
    st.markdown("### 📁 取证预览")
    st.markdown("上传压缩包，预览并勾选需要识别的截图")

    uploaded_file = st.file_uploader(
        "上传取证压缩包",
        type=['zip'],
        help="支持 .zip 格式"
    )

    if uploaded_file:
        # 保存上传的文件
        zip_path = TEMP_DIR / uploaded_file.name
        with open(zip_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"✅ 文件已上传：{uploaded_file.name}")

        # 解压
        filename = Path(uploaded_file.name)
        extract_dir = TEMP_DIR / filename.stem
        extract_dir.mkdir(exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        st.success(f"✅ 解压完成！")

        # 获取所有图片文件
        image_files = []
        for ext in ['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp']:
            image_files.extend(extract_dir.rglob(ext))

        # 排除录屏文件夹
        image_files = [f for f in image_files if 'screen' not in f.name.lower() and 'record' not in f.name.lower()]

        st.info(f"📊 共找到 {len(image_files)} 张图片")

        # 初始化 session state
        if 'selected_images' not in st.session_state:
            st.session_state.selected_images = []

        # 初始化页码
        if 'page' not in st.session_state:
            st.session_state.page = 1

        # 显示选中数量（顶部）
        st.markdown(f"**已选择：{len(st.session_state.selected_images)} / {len(image_files)} 张**")

        # 分页显示（24张/页）
        PAGE_SIZE = 24  # 每页 4 列 × 6 行
        total_pages = (len(image_files) + PAGE_SIZE - 1) // PAGE_SIZE
        page = st.session_state.page

        start_idx = (page - 1) * PAGE_SIZE
        end_idx = min(start_idx + PAGE_SIZE, len(image_files))

        # 缓存缩略图（带 hash 缓存）
        @st.cache_data
        def load_thumbnail(img_path_str):
            img_path = Path(img_path_str)
            try:
                img = Image.open(img_path)
                img.thumbnail((240, 180))
                return img
            except:
                return None

        # 显示图片网格
        cols = st.columns(4)
        for idx, img_path in enumerate(image_files[start_idx:end_idx]):
            col_idx = idx % 4
            with cols[col_idx]:
                try:
                    img = load_thumbnail(str(img_path))
                    if img is None:
                        continue

                    relative_path = str(img_path.relative_to(extract_dir))
                    is_selected = relative_path in st.session_state.selected_images

                    # 显示图片
                    st.image(img, use_container_width=True)

                    # 显示文件名
                    st.caption(img_path.name[:15] + "...")

                    # 选择按钮
                    button_key = f"select_{idx}_{page}"
                    if st.button(
                        "✓" if is_selected else "○",
                        key=button_key,
                        use_container_width=True
                    ):
                        if is_selected:
                            st.session_state.selected_images.remove(relative_path)
                        else:
                            st.session_state.selected_images.append(relative_path)
                        st.rerun()

                except Exception as e:
                    st.error("加载失败")

        # 分页导航和操作按钮（底部）
        st.markdown("---")
        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
        with col1:
            select_all = st.button("✅ 全选")
        with col2:
            clear_all = st.button("❌ 清空")
        with col3:
            if st.button("⬅️ 上一页", disabled=page <= 1):
                st.session_state.page = page - 1
                st.rerun()
        with col4:
            if st.button("下一页 ➡️", disabled=page >= total_pages):
                st.session_state.page = page + 1
                st.rerun()

        # 显示当前页码
        st.markdown(f"<p style='text-align:center; color:#666;'>第 {page} / {total_pages} 页</p>", unsafe_allow_html=True)

        # 处理全选/清空
        if select_all:
            st.session_state.selected_images = [str(f.relative_to(extract_dir)) for f in image_files]
            st.rerun()

        if clear_all:
            st.session_state.selected_images = []
            st.rerun()

        # 导出按钮
        st.markdown("---")
        if st.button("📥 导出选中图片", type="primary"):
            if st.session_state.selected_images:
                df = pd.DataFrame({
                    '序号': range(1, len(st.session_state.selected_images) + 1),
                    '图片路径': st.session_state.selected_images,
                    '来源压缩包': [uploaded_file.name] * len(st.session_state.selected_images),
                    '完整路径': [str(extract_dir / path) for path in st.session_state.selected_images]
                })

                csv_path = TEMP_DIR / "selected_images.csv"
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')

                with open(csv_path, 'rb') as f:
                    st.download_button(
                        label="⬇️ 下载 CSV 文件",
                        data=f,
                        file_name="selected_images.csv",
                        mime="text/csv"
                    )

                st.success(f"✅ 已导出 {len(st.session_state.selected_images)} 张图片！")
            else:
                st.warning("⚠️ 请先勾选至少一张图片")

        # 保存到 session state，供 Tab 2 使用
        st.session_state.extract_dir = str(extract_dir)
        st.session_state.zip_name = uploaded_file.name


# ============ Tab 2: OCR 识别 ============
with tab2:
    st.markdown("### 🔍 OCR 识别")
    st.markdown("批量识别截图中的单号和日期")

    # 检查是否有选中图片
    if 'selected_images' in st.session_state and st.session_state.selected_images:
        st.info(f"📋 已选择 {len(st.session_state.selected_images)} 张图片待识别")

        # 调试模式选项
        debug_mode = st.checkbox("🔧 调试模式（显示OCR识别的原始文本）", help="开启后可以看到每张图片的OCR识别原始文本，帮助诊断识别问题")

        if st.button("🚀 开始 OCR 识别", type="primary"):
            extract_dir = Path(st.session_state.extract_dir)
            zip_name = st.session_state.zip_name

            progress_bar = st.progress(0)
            status_text = st.empty()

            all_records = []
            debug_info = []  # 存储调试信息

            for idx, img_path in enumerate(st.session_state.selected_images):
                progress = (idx + 1) / len(st.session_state.selected_images)
                progress_bar.progress(progress)
                status_text.text(f"正在识别 {idx + 1} / {len(st.session_state.selected_images)}: {img_path}")

                try:
                    full_path = extract_dir / img_path
                    img = Image.open(full_path)

                    # OCR 识别
                    text = pytesseract.image_to_string(img, lang='chi_sim+eng')

                    # 提取截图名称
                    截图名称 = Path(img_path).name

                    # 提取反馈产品（更灵活的正则）
                    反馈产品 = "未识别"
                    # 尝试多种格式
                    product_patterns = [
                        r'反馈产品\s*[:：]?\s*([^\n]+)',  # 反馈产品: xxx 或 反馈产品：xxx
                        r'反馈产品[^\n]*\n([^\n]+)',          # 反馈产品后换行
                        r'产品\s*[:：]?\s*([^\n]+)'            # 简化版：产品
                    ]
                    for pattern in product_patterns:
                        product_match = re.search(pattern, text)
                        if product_match:
                            product_text = product_match.group(1).strip()
                            # 清理多余字符，保留中文、英文、数字
                            product_text = re.sub(r'[^\w\u4e00-\u9fff\s]', '', product_text)
                            # 移除多余空格
                            product_text = ' '.join(product_text.split())
                            if product_text and len(product_text) > 1:
                                反馈产品 = product_text
                                break

                    # 提取作品名称（更灵活的正则）
                    作品名称 = "未识别"
                    work_patterns = [
                        r'作品名称\s*[:：]?\s*([^\n]+)',  # 作品名称: xxx
                        r'作品名称[^\n]*\n([^\n]+)',          # 作品名称后换行
                        r'作品\s*[:：]?\s*([^\n]+)'            # 简化版：作品
                    ]
                    for pattern in work_patterns:
                        work_match = re.search(pattern, text)
                        if work_match:
                            work_text = work_match.group(1).strip()
                            work_text = re.sub(r'[^\w\u4e00-\u9fff\s]', '', work_text)
                            work_text = ' '.join(work_text.split())
                            if work_text and len(work_text) > 1:
                                作品名称 = work_text
                                break

                    # 提取单号（8位数字）
                    numbers = re.findall(r'\d{8}', text)

                    # 提取日期（格式：20xx-xx-xx）
                    dates = re.findall(r'20\d{2}-\d{2}-\d{2}', text)

                    # 保存调试信息
                    if debug_mode:
                        debug_info.append({
                            '截图名称': 截图名称,
                            'OCR文本': text,
                            '反馈产品': 反馈产品,
                            '作品名称': 作品名称,
                            '识别单号': numbers,
                            '识别日期': dates
                        })

                    # 匹配单号和日期
                    if numbers and dates:
                        min_len = min(len(numbers), len(dates))
                        for i in range(min_len):
                            all_records.append({
                                '截图名称': 截图名称,
                                '反馈产品': 反馈产品,
                                '作品名称': 作品名称,
                                '单号': numbers[i],
                                '日期': dates[i],
                                '来源压缩包': zip_name
                            })

                except Exception as e:
                    st.error(f"❌ 识别失败 {img_path}: {e}")
                    if debug_mode:
                        debug_info.append({
                            '截图名称': img_path,
                            'OCR文本': f"识别失败: {str(e)}",
                            '反馈产品': "失败",
                            '作品名称': "失败",
                            '识别单号': [],
                            '识别日期': []
                        })

            progress_bar.progress(1.0)
            status_text.text("✅ 识别完成！")

            if all_records:
                # 转换为 DataFrame
                result_df = pd.DataFrame(all_records)

                # 去重
                result_df = result_df.drop_duplicates(subset=['单号', '日期'])

                # 导出 Excel
                output_path = TEMP_DIR / "ocr_results.xlsx"
                result_df.to_excel(output_path, index=False, engine='openpyxl')

                # 显示统计
                st.success(f"✅ 识别完成！共 {len(result_df)} 条记录")

                st.markdown("### 📋 识别结果预览")
                st.dataframe(result_df.head(10), use_container_width=True)

                # 显示调试信息（如果开启）
                if debug_mode and debug_info:
                    st.markdown("---")
                    st.markdown("### 🔧 调试信息（OCR识别原始文本）")
                    for info in debug_info:
                        with st.expander(f"📸 {info['截图名称']}", expanded=False):
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown(f"**反馈产品：** {info['反馈产品']}")
                                st.markdown(f"**作品名称：** {info['作品名称']}")
                            with col2:
                                st.markdown(f"**识别单号：** {', '.join(info['识别单号']) if info['识别单号'] else '无'}")
                                st.markdown(f"**识别日期：** {', '.join(info['识别日期']) if info['识别日期'] else '无'}")

                            st.markdown("**OCR原始文本：**")
                            st.code(info['OCR文本'], language='text')

                # 提供下载
                with open(output_path, 'rb') as f:
                    st.download_button(
                        label="⬇️ 下载 OCR 结果 (Excel)",
                        data=f,
                        file_name="ocr_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # 保存到 session state，供 Tab 3 使用
                st.session_state.ocr_results = result_df
            else:
                st.warning("⚠️ 未识别到任何记录，请检查图片清晰度")
    else:
        st.warning("⚠️ 请先在「取证预览」标签页中勾选图片")


# ============ Tab 3: Excel 匹配 ============
with tab3:
    st.markdown("### 📊 Excel 匹配")
    st.markdown("匹配总表和 OCR 结果，标记取证来源")

    # 检查是否有 OCR 结果
    if 'ocr_results' not in st.session_state:
        st.warning("⚠️ 请先在「OCR 识别」标签页中完成识别")

        # 也支持手动上传
        st.markdown("---")
        st.markdown("#### 或手动上传 OCR 结果")
        manual_ocr_file = st.file_uploader(
            "上传 OCR 结果 Excel 文件",
            type=['xlsx'],
            help="或使用「OCR 识别」标签页自动生成的文件"
        )

        if manual_ocr_file:
            st.session_state.ocr_results = pd.read_excel(manual_ocr_file)
            st.success(f"✅ 已加载 {len(st.session_state.ocr_results)} 条 OCR 记录")
    else:
        st.success(f"✅ 已加载 {len(st.session_state.ocr_results)} 条 OCR 记录")

    # 上传总表
    st.markdown("---")
    master_file = st.file_uploader(
        "上传总表 Excel 文件",
        type=['xlsx'],
        help="包含投诉单号和日期的总表"
    )

    if master_file and 'ocr_results' in st.session_state:
        st.success(f"✅ 总表已上传")

        # 读取总表
        master_df = pd.read_excel(master_file)

        st.markdown("#### 📋 总表列名匹配")
        st.write("可用列名：", list(master_df.columns))

        # 只选择单号列
        id_col = st.selectbox(
            "选择投诉单号列",
            options=list(master_df.columns),
            index=list(master_df.columns).index('投诉单号') if '投诉单号' in master_df.columns else 0
        )

        # 开始匹配
        if st.button("🔍 开始匹配", type="primary"):
            st.info("正在匹配...")

            # 创建 OCR 查找字典（按单号匹配）
            ocr_df = st.session_state.ocr_results
            ocr_dict = {}
            ocr_date_dict = {}

            for _, row in ocr_df.iterrows():
                key = str(row['单号'])
                # 取证来源：来源压缩包 - 截图名称
                ocr_dict[key] = f"{row['来源压缩包']} - {row['截图名称']}"
                # 识别日期
                ocr_date_dict[key] = row['日期']

            # 匹配总表
            master_df['取证来源'] = '未匹配'
            master_df['识别日期'] = ''
            matched_count = 0

            for idx, row in master_df.iterrows():
                key = str(row[id_col])
                if key in ocr_dict:
                    master_df.at[idx, '取证来源'] = ocr_dict[key]
                    master_df.at[idx, '识别日期'] = ocr_date_dict[key]
                    matched_count += 1

            # 导出结果
            output_path = TEMP_DIR / "匹配结果.xlsx"
            master_df.to_excel(output_path, index=False, engine='openpyxl')

            # 显示统计
            st.success(f"✅ 匹配完成！")
            st.markdown(f"""
            ### 📊 匹配统计
            - 总表记录数：{len(master_df)}
            - ✅ 成功匹配：{matched_count} 条 ({matched_count/len(master_df)*100:.1f}%)
            - ❌ 未匹配：{len(master_df) - matched_count} 条
            """)

            # 显示预览
            st.markdown("#### 📋 匹配结果预览")
            st.dataframe(master_df.head(10), use_container_width=True)

            # 提供下载
            with open(output_path, 'rb') as f:
                st.download_button(
                    label="⬇️ 下载匹配结果 (Excel)",
                    data=f,
                    file_name="匹配结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 导出未匹配列表
            unmatched_df = master_df[master_df['取证来源'] == '未匹配']
            if not unmatched_df.empty:
                unmatched_path = TEMP_DIR / "未匹配列表.xlsx"
                unmatched_df.to_excel(unmatched_path, index=False, engine='openpyxl')

                with open(unmatched_path, 'rb') as f:
                    st.download_button(
                        label="⬇️ 下载未匹配列表 (Excel)",
                        data=f,
                        file_name="未匹配列表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )


# 侧边栏清理按钮
st.sidebar.markdown("---")
if st.sidebar.button("🗑️ 清理临时文件"):
    import shutil
    try:
        shutil.rmtree(TEMP_DIR)
        TEMP_DIR.mkdir(exist_ok=True)
        st.session_state.selected_images = []
        st.session_state.ocr_results = None
        st.success("✅ 临时文件已清理，请刷新页面")
    except Exception as e:
        st.error(f"清理失败：{e}")
