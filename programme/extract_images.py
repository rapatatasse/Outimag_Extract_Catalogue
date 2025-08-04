import os
import fitz  # PyMuPDF
from PIL import Image
import io

def extract_and_convert_images(pdf_path, output_dir):
    """Extracts images from a PDF, converts them to PNG, and saves them."""
    try:
        doc = fitz.open(pdf_path)
        pdf_filename_base = os.path.splitext(os.path.basename(pdf_path))[0]
        image_count_for_pdf = 0

        for page_num in range(len(doc)):
            for img_index, img in enumerate(doc.get_page_images(page_num, full=True)):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]

                # Convert to PNG using Pillow
                try:
                    image = Image.open(io.BytesIO(image_bytes))
                    image_count_for_pdf += 1
                    
                    # New filename format: R12345-1.png
                    new_filename = f"{pdf_filename_base}-{image_count_for_pdf}.png"
                    output_path = os.path.join(output_dir, new_filename)
                    
                    image.save(output_path, "PNG")
                except Exception as img_e:
                    print(f"Could not process image {img_index} on page {page_num+1} in {os.path.basename(pdf_path)}: {img_e}")

        doc.close()
        if image_count_for_pdf > 0:
            print(f"Successfully extracted and converted {image_count_for_pdf} images from {os.path.basename(pdf_path)}")
        else:
            print(f"No images found in {os.path.basename(pdf_path)}")

    except Exception as e:
        print(f"Error processing file {os.path.basename(pdf_path)}: {e}")

def main():
    """Main function to process PDFs starting with 'R'."""
    pdf_directory = os.path.join(os.path.dirname(__file__), './A Fiches techniques a traiter')
    output_directory = os.path.join(os.path.dirname(__file__), './extracted_images')

    if not os.path.isdir(pdf_directory):
        print(f"Error: Directory not found at '{pdf_directory}'")
        return
        
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    print(f"Searching for PDF files starting with 'R' in: {pdf_directory}")
    print(f"Saving extracted images to: {output_directory}\n")

    for filename in os.listdir(pdf_directory):
        # Only process PDFs starting with 'R' (case-insensitive)
        if filename.lower().startswith('r') and filename.lower().endswith('.pdf'):
            print(f"--- Processing file: {filename} ---")
            pdf_path = os.path.join(pdf_directory, filename)
            extract_and_convert_images(pdf_path, output_directory)
            print("-" * (len(filename) + 22) + "\n")

    print("\nImage extraction complete.")

if __name__ == "__main__":
    main()
