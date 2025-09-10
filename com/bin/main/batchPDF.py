import fitz  # PyMuPDF
import os
import time
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# 设置本地 Tesseract 可执行文件路径
pytesseract.pytesseract.tesseract_cmd = r"E:\OCR\tesseract.exe"
# 设置 Poppler bin 路径
POPPLER_PATH = r'E:\poppler\Release-25.07.0-0\poppler-25.07.0\Library\bin'


def pdf_to_md(pdf_path, output_md, image_dir="../../images", ocr_threshold=20):
    os.makedirs(image_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    md_lines = []

    total_pages = doc.page_count
    start_time = time.time()
    print(f"\n开始处理 {pdf_path}，总页数 {total_pages}")

    for page_num, page in enumerate(doc, start=1):
        page_start = time.time()
        print(f"处理第 {page_num}/{total_pages} 页...")

        text = page.get_text().strip()

        # 如果文字太少，可能是扫描件，转为图片并 OCR
        if len(text) < ocr_threshold:
            images = convert_from_path(
                pdf_path,
                first_page=page_num,
                last_page=page_num,
                poppler_path=POPPLER_PATH
            )
            ocr_text = ""
            for i, img in enumerate(images, start=1):
                img_path = os.path.join(image_dir, f"page_{page_num}_img_{i}.png")
                img.save(img_path, "PNG")
                print(f"  OCR 第 {i} 张图片...")
                ocr_text += pytesseract.image_to_string(img, lang="chi_sim+eng")
                md_lines.append(f"![page_{page_num}_img_{i}]({img_path})")
            if ocr_text.strip():
                md_lines.append(ocr_text.strip())
        else:
            # 提取页面图片
            for img_index, img in enumerate(page.get_images(full=True), start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                img_bytes = base_image["image"]
                img_ext = base_image["ext"]
                img_path = os.path.join(image_dir, f"page_{page_num}_img_{img_index}.{img_ext}")
                with open(img_path, "wb") as f:
                    f.write(img_bytes)
                md_lines.append(f"![page_{page_num}_img_{img_index}]({img_path})")
            md_lines.append(text)

        # 每页处理完成后打印 ETA
        elapsed = time.time() - start_time
        pages_done = page_num
        pages_left = total_pages - pages_done
        eta = elapsed / pages_done * pages_left
        print(f"已用时间 {elapsed:.1f}s，预计剩余 {eta:.1f}s")

    with open(output_md, "w", encoding="utf-8") as f:
        f.write("\n\n".join(md_lines))
    print(f"完成 {pdf_path} → {output_md}")


def batch_convert(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    total_files = len(pdf_files)
    start_time_all = time.time()
    print(f"\n找到 {total_files} 个 PDF 文件，开始批量转换...")

    for idx, file in enumerate(pdf_files, start=1):
        pdf_path = os.path.join(folder_path, file)
        base_name = os.path.splitext(file)[0]
        output_dir = os.path.join(folder_path, base_name)
        os.makedirs(output_dir, exist_ok=True)
        output_md = os.path.join(output_dir, f"{base_name}.md")
        image_dir = os.path.join(output_dir, "images")
        print(f"\n处理第 {idx}/{total_files} 个文件：{file}")
        pdf_to_md(pdf_path, output_md, image_dir=image_dir)

        # 打印总 ETA
        elapsed_all = time.time() - start_time_all
        files_done = idx
        files_left = total_files - files_done
        eta_all = elapsed_all / files_done * files_left
        print(f"已处理 {files_done}/{total_files} 个文件，总已用 {elapsed_all:.1f}s，预计总剩余 {eta_all:.1f}s")

    print("\n全部 PDF 文件处理完成！")


if __name__ == "__main__":
    folder = r"../../files"  # 替换成你的 PDF 文件夹路径
    batch_convert(folder)
