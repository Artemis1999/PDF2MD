import fitz  # PyMuPDF
import os
from pdf2image import convert_from_path
import pytesseract
from PIL import Image

# 如果你使用本地 Tesseract，需要设置路径
pytesseract.pytesseract.tesseract_cmd = r"E:\OCR\tesseract.exe"


def pdf_to_md(pdf_path, output_md, image_dir="images", ocr_threshold=20):
    os.makedirs(image_dir, exist_ok=True)
    doc = fitz.open(pdf_path)

    md_lines = []
    for page_num, page in enumerate(doc, start=1):
        text = page.get_text().strip()

        # 如果文字太少，可能是扫描件，转为图片并 OCR
        if len(text) < ocr_threshold:
            images = convert_from_path(pdf_path, first_page=page_num, last_page=page_num,poppler_path=r'E:\poppler\Release-25.07.0-0\poppler-25.07.0\Library\bin')
            ocr_text = ""
            for i, img in enumerate(images, start=1):
                img_path = os.path.join(image_dir, f"page_{page_num}_img_{i}.png")
                img.save(img_path, "PNG")
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

    with open(output_md, "w", encoding="utf-8") as f:
        f.write("\n\n".join(md_lines))


if __name__ == "__main__":
    pdf_to_md("../../files/sample.pdf", "../../output/sample.md")
