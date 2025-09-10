import fitz
import os
import time
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from multiprocessing import Pool, cpu_count
import camelot  # 新增表格识别

# 设置 Tesseract 和 Poppler 路径
pytesseract.pytesseract.tesseract_cmd = r"E:\OCR\tesseract.exe"
POPPLER_PATH = r"E:\poppler\Release-25.07.0-0\poppler-25.07.0\Library\bin"

OCR_THRESHOLD = 20  # 文本长度阈值

# 工程图/表格页默认当图片处理
SCAN_AS_IMAGE_PAGES = []  # 可以设置特定页号，或者留空默认按 OCR 阈值判断


def pdf_to_md_single(pdf_info):
    pdf_path, output_dir = pdf_info
    os.makedirs(output_dir, exist_ok=True)
    image_dir = os.path.join(output_dir, "images")
    os.makedirs(image_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_md = os.path.join(output_dir, f"{base_name}.md")

    doc = fitz.open(pdf_path)
    md_lines = []
    total_pages = doc.page_count
    start_time = time.time()

    print(f"\n开始处理 {pdf_path}，总页数 {total_pages}")

    for page_num, page in enumerate(doc, start=1):
        page_start = time.time()
        print(f"处理第 {page_num}/{total_pages} 页...")

        text = page.get_text().strip()

        # 判断扫描页/工程图页
        if len(text) < OCR_THRESHOLD or page_num in SCAN_AS_IMAGE_PAGES:
            images = convert_from_path(
                pdf_path, first_page=page_num, last_page=page_num, poppler_path=POPPLER_PATH
            )
            ocr_text = ""
            for i, img in enumerate(images, start=1):
                img_path = os.path.join(image_dir, f"page_{page_num}_img_{i}.png")
                img.save(img_path, "PNG")
                # OCR 识别文字
                ocr_text += pytesseract.image_to_string(img, lang="chi_sim+eng")
                md_lines.append(f"![page_{page_num}_img_{i}]({img_path})")
            if ocr_text.strip():
                md_lines.append(ocr_text.strip())
        else:
            # 尝试用 Camelot 提取表格
            try:
                tables = camelot.read_pdf(pdf_path, pages=str(page_num))
                if tables:
                    for i, table in enumerate(tables):
                        md_table = table.df.to_markdown()
                        md_lines.append(md_table)
                    continue  # 如果有表格，跳过文本/图片提取
            except Exception as e:
                print(f"第 {page_num} 页表格提取失败: {e}")

            # 普通文字
            md_lines.append(text)

            # 页面图片
            for img_index, img in enumerate(page.get_images(full=True), start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                img_bytes = base_image["image"]
                img_ext = base_image["ext"]
                img_path = os.path.join(image_dir, f"page_{page_num}_img_{img_index}.{img_ext}")
                with open(img_path, "wb") as f:
                    f.write(img_bytes)
                md_lines.append(f"![page_{page_num}_img_{img_index}]({img_path})")

        elapsed = time.time() - start_time
        eta = elapsed / page_num * (total_pages - page_num)
        print(f"已用时间 {elapsed:.1f}s，预计剩余 {eta:.1f}s")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write("\n\n".join(md_lines))

    print(f"完成 {pdf_path} → {output_md}")


def batch_convert(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    total_files = len(pdf_files)
    print(f"\n找到 {total_files} 个 PDF 文件，开始批量转换...")

    pdf_info_list = [
        (os.path.join(folder_path, pdf), os.path.join(folder_path, os.path.splitext(pdf)[0]))
        for pdf in pdf_files
    ]

    start_time_all = time.time()
    # 多进程
    with Pool(min(cpu_count(), total_files)) as pool:
        pool.map(pdf_to_md_single, pdf_info_list)

    elapsed_all = time.time() - start_time_all
    print(f"\n全部 PDF 文件处理完成，总耗时 {elapsed_all:.1f}s")


if __name__ == "__main__":
    folder = r"../../files"  # PDF 文件夹路径
    batch_convert(folder)
