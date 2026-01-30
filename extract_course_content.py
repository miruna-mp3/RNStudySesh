from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pdfplumber
from PIL import Image
import os
import io

def extract_from_pptx(file_path, course_num):
    """Extract text and images from PowerPoint file"""
    try:
        prs = Presentation(file_path)
        text_content = []
        image_count = 0
        
        # Create image directory for this course
        img_dir = f"images/course{course_num}"
        os.makedirs(img_dir, exist_ok=True)
        
        for slide_num, slide in enumerate(prs.slides, 1):
            text_content.append(f"\n{'─' * 60}")
            text_content.append(f"SLIDE {slide_num}")
            text_content.append(f"{'─' * 60}\n")
            
            slide_images = []
            slide_text = []
            
            for shape in slide.shapes:
                # Extract text
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text.strip())
                
                # Extract images
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image_count += 1
                    try:
                        image = shape.image
                        image_bytes = image.blob
                        img_filename = f"slide{slide_num}_img{image_count}.{image.ext}"
                        img_path = os.path.join(img_dir, img_filename)
                        
                        with open(img_path, 'wb') as img_file:
                            img_file.write(image_bytes)
                        
                        slide_images.append(img_filename)
                    except Exception as e:
                        print(f"Error extracting image: {e}")
            
            # Add text content
            if slide_text:
                text_content.append("\n".join(slide_text))
            else:
                text_content.append("[No text on this slide]")
            
            # Add image references
            if slide_images:
                text_content.append(f"\n📷 Images on this slide:")
                for img in slide_images:
                    text_content.append(f"   - {img_dir}/{img}")
            
            text_content.append("")
        
        return "\n".join(text_content), image_count
    except Exception as e:
        return f"Error extracting from {file_path}: {str(e)}", 0

def extract_from_pdf(file_path, course_num):
    """Extract text and images from PDF file"""
    try:
        text_content = []
        image_count = 0
        
        # Create image directory for this course
        img_dir = f"images/course{course_num}"
        os.makedirs(img_dir, exist_ok=True)
        
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text_content.append(f"\n{'─' * 60}")
                text_content.append(f"PAGE {page_num}")
                text_content.append(f"{'─' * 60}\n")
                
                # Extract text
                text = page.extract_text()
                if text:
                    text_content.append(text.strip())
                else:
                    text_content.append("[No text on this page]")
                
                # Extract images
                page_images = []
                if hasattr(page, 'images'):
                    for img_idx, img_info in enumerate(page.images, 1):
                        image_count += 1
                        try:
                            img_filename = f"page{page_num}_img{img_idx}.png"
                            img_path = os.path.join(img_dir, img_filename)
                            
                            # Extract image using pdfplumber
                            x0, y0, x1, y1 = img_info['x0'], img_info['top'], img_info['x1'], img_info['bottom']
                            cropped = page.crop((x0, y0, x1, y1))
                            img = cropped.to_image(resolution=150)
                            img.save(img_path)
                            
                            page_images.append(img_filename)
                        except Exception as e:
                            print(f"Error extracting image from page {page_num}: {e}")
                
                if page_images:
                    text_content.append(f"\n📷 Images on this page:")
                    for img in page_images:
                        text_content.append(f"   - {img_dir}/{img}")
                
                text_content.append("")
        
        return "\n".join(text_content), image_count
    except Exception as e:
        return f"Error extracting from {file_path}: {str(e)}", 0

def main():
    # Define course files in order
    course_files = [
        ("curs1 RN 2025.pptx", "Course 1: RN 2025"),
        ("curs2 RN 2025 - perceptron.pptx", "Course 2: Perceptron"),
        ("curs3 RN 2025 -  gradient descent.pptx", "Course 3: Gradient Descent"),
        ("curs4 RN 2025 - backpropagation.pptx", "Course 4: Backpropagation"),
        ("curs5 - weight initialization & overfitting.pptx", "Course 5: Weight Initialization & Overfitting"),
        ("curs6 - optimizers.pdf", "Course 6: Optimizers"),
        ("Curs 7 rn - pytorch.pptx", "Course 7: PyTorch"),
        ("curs8 RN - Q Learning.pptx", "Course 8: Q Learning"),
        ("curs9 -  Convolutional.pptx", "Course 9: Convolutional Networks"),
        ("curs10 - actor critic.pptx", "Course 10: Actor Critic"),
        ("curs11 - LSTM.pptx", "Course 11: LSTM"),
    ]
    
    all_content = []
    total_images = 0
    
    # Create main images directory
    os.makedirs("images", exist_ok=True)
    
    for course_num, (file_name, course_title) in enumerate(course_files, 1):
        if not os.path.exists(file_name):
            print(f"Warning: {file_name} not found, skipping...")
            continue
        
        print(f"Extracting {course_title}...")
        all_content.append("\n\n")
        all_content.append("=" * 80)
        all_content.append(course_title.upper().center(80))
        all_content.append("=" * 80)
        
        if file_name.endswith('.pptx'):
            content, img_count = extract_from_pptx(file_name, course_num)
        elif file_name.endswith('.pdf'):
            content, img_count = extract_from_pdf(file_name, course_num)
        else:
            content = f"Unknown file format: {file_name}"
            img_count = 0
        
        all_content.append(content)
        total_images += img_count
        print(f"  Extracted {img_count} images")
    
    # Write all content to a single file
    output_file = "course_content.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("\n".join(all_content))
    
    print(f"\n✅ Content extracted successfully!")
    print(f"   - Text saved to: {output_file}")
    print(f"   - Total images extracted: {total_images}")
    print(f"   - Images saved in: images/course1/ through images/course11/")

if __name__ == "__main__":
    main()