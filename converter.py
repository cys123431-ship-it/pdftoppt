import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import os

def convert_pdf_to_pptx(pdf_path, pptx_path, progress_callback=None):
    """
    Converts a PDF file to a PowerPoint presentation where each page is an image.

    Args:
        pdf_path (str): Path to the source PDF file.
        pptx_path (str): Path to save the resulting PPTX file.
        progress_callback (function): Optional callback function that takes a percentage (0-100).
    """
    try:
        doc = fitz.open(pdf_path)
        prs = Presentation()

        # Set slide dimensions to match the first page of the PDF (assuming uniform size)
        # Defaults to standard 4:3 if PDF is empty
        if len(doc) > 0:
            page = doc[0]
            # PDF points are 1/72 inch, PPTX uses EMU (English Metric Unit)
            # simplest is to convert points to inches then let python-pptx handle conversion internally via Inches()
            # fitz.rect width/height are in points
            width_inches = page.rect.width / 72.0
            height_inches = page.rect.height / 72.0
            
            prs.slide_width = Inches(width_inches)
            prs.slide_height = Inches(height_inches)

        total_pages = len(doc)
        
        for i, page in enumerate(doc):
            # Render page to image (PixMap)
            # zoom_x and zoom_y set the resolution. 2.0 means 144 DPI (2 * 72)
            # Higher values mean better quality but larger file size
            zoom = 2.0 
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to image bytes
            img_data = pix.tobytes("png")
            image_stream = io.BytesIO(img_data)
            
            # Add a blank slide
            # 6 is the index for a blank layout in the default template
            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # Add image to slide, filling the entire slide
            slide.shapes.add_picture(
                image_stream, 
                0, 
                0, 
                width=prs.slide_width, 
                height=prs.slide_height
            )

            # Update progress
            if progress_callback:
                percent = int((i + 1) / total_pages * 100)
                progress_callback(percent)

        prs.save(pptx_path)
        return True, "Conversion successful!"

    except Exception as e:
        return False, str(e)
