# ocr.py

import os
import time
from datetime import datetime
from dotenv import load_dotenv
from paddleocr import PaddleOCR
from pdf2image import convert_from_path  # still available if you ever need PDF → images
from llm import run_llm_table_extraction  # import from llm.py

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))

# If you still ever use PDF → image on Windows:
poppler_path = r"C:\poppler-24.08.0\Library\bin"

# Initialize PaddleOCR once
ocr = PaddleOCR(use_angle_cls=True, lang="en")


# ---------------------------------------------------------
# Optional: PDF → images (kept as utility if you still need it)
# ---------------------------------------------------------
def pdf_to_images(pdf_path, output_folder, poppler_path=poppler_path, dpi=300, image_format="png"):
    os.makedirs(output_folder, exist_ok=True)
    print(f"Converting '{pdf_path}' to images...")

    pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)

    image_paths = []
    for idx, page in enumerate(pages, start=1):
        image_filename = os.path.join(output_folder, f"page_{idx}.{image_format}")
        page.save(image_filename, image_format.upper())
        image_paths.append(image_filename)
        print(f"Saved: {image_filename}")

    print(f"PDF successfully converted! Images saved in: {output_folder}")
    return image_paths


# ---------------------------------------------------------
# OCR on a single image → write extracted_text.txt (Option A: overwrite)
# ---------------------------------------------------------
def extract_text_from_image(image_path, output_file="extracted_text.txt"):
    """
    Runs OCR on a single image and writes ONLY this image's OCR text
    into output_file (overwrites previous content).
    """
    try:
        if not os.path.exists(image_path):
            print(f"Error: Image file '{image_path}' not found!")
            return None

        print(f"\n📸 Processing image: {image_path}")

        result = ocr.predict(image_path)
        extracted_text_lines = []

        if result and len(result) > 0:
            for line in result:
                try:
                    if isinstance(line, list) and len(line) >= 2:
                        if isinstance(line[1], list) and len(line[1]) >= 2:
                            text = line[1][0]
                            confidence = line[1][1]
                        elif isinstance(line[1], str):
                            text = line[1]
                            confidence = 1.0
                        else:
                            text = str(line[1])
                            confidence = 1.0
                    elif isinstance(line, tuple) and len(line) >= 2:
                        text = line[1]
                        confidence = 1.0
                    else:
                        text = str(line)
                        confidence = 1.0

                    extracted_text_lines.append(f"{text} (confidence: {confidence:.2f})")
                except Exception as e:
                    print(f"Error processing line {line}: {e}")
                    extracted_text_lines.append(f"Error processing line: {str(line)}")

        # Ensure folder for output file exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # 🔴 OPTION A: overwrite previous OCR text completely
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n" + "=" * 70 + "\n")
            f.write(f"Text extracted from: {image_path}\n")
            f.write(f"Extraction time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("-" * 50 + "\n\n")

            if extracted_text_lines:
                for text in extracted_text_lines:
                    f.write(text + "\n")
            else:
                f.write("No text found in the image.\n")

        print(f"✅ Text extraction completed for {image_path}")
        return result

    except Exception as e:
        print(f"Error processing image {image_path}: {str(e)}")
        return None


# ---------------------------------------------------------
# NEW: Process a folder of images, send each image to LLM
# ---------------------------------------------------------
def process_images_folder(
    images_folder: str,
    output_text_file: str = "extracted_text.txt",
    output_xlsx: str = "table_output.xlsx",
):
    """
    Main new flow:

    - Loop over all images in `images_folder`
    - For each image:
        * Run OCR → overwrite `output_text_file` with this image's OCR
        * Call Gemini via run_llm_table_extraction()
        * Append resulting table rows to `output_xlsx`
    """

    if not os.path.isdir(images_folder):
        raise NotADirectoryError(f"❌ Folder does not exist: {images_folder}")

    # Supported image extensions
    exts = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}
    all_files = sorted(os.listdir(images_folder))

    image_paths = [
        os.path.join(images_folder, f)
        for f in all_files
        if os.path.splitext(f.lower())[1] in exts
    ]

    if not image_paths:
        print(f"⚠️ No image files found in: {images_folder}")
        return

    print(f"Found {len(image_paths)} image(s) to process in '{images_folder}'.")

    for idx, img_path in enumerate(image_paths, start=1):
        print(f"\n===== [{idx}/{len(image_paths)}] START IMAGE =====")

        # 1) OCR → extracted_text.txt (overwrite)
        extract_text_from_image(img_path, output_file=output_text_file)

        # 2) LLM → table + metadata → append to Excel
        run_llm_table_extraction(
            input_file=output_text_file,
            output_xlsx=output_xlsx,
        )

        print(f"===== [{idx}/{len(image_paths)}] DONE IMAGE =====\n")
        # small delay if you want to be nice to API rate limits
        time.sleep(1)


if __name__ == "__main__":
    # 🔧 CHANGE THIS to your folder containing invoice screenshots
    images_folder = r"invoice_images"   # e.g. "AmanRMC_Images" or full path

    process_images_folder(
        images_folder=images_folder,
        output_text_file="extracted_text.txt",
        output_xlsx="table_output.xlsx",
    )
