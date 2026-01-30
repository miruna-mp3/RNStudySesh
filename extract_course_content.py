from pptx import Presentation
import pdfplumber
import os

def extract_text_from_pptx(file_path):
    """Extract text from PowerPoint file"""
    try:
        prs = Presentation(file_path)
        text_content = []
        
        for slide_num, slide in enumerate(prs.slides, 1):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    if shape.text.strip():
                        slide_text.append(shape.text)
            
            if slide_text:
                text_content.append(f"Slide {slide_num}:\n" + "\n".join(slide_text))
        
        return "\n\n".join(text_content)
    except Exception as e:
        return f"Error extracting from {file_path}: {str(e)}"

def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    try:
        text_content = []
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    text_content.append(f"Page {page_num}:\n{text}")
        
        return "\n\n".join(text_content)
    except Exception as e:
        return f"Error extracting from {file_path}: {str(e)}"

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
    
    for file_name, course_title in course_files:
        if not os.path.exists(file_name):
            print(f"Warning: {file_name} not found, skipping...")
            continue
        
        print(f"Extracting {course_title}...")
        all_content.append("=" * 80)
        all_content.append(course_title.upper())
        all_content.append("=" * 80)
        all_content.append("")
        
        if file_name.endswith('.pptx'):
            content = extract_text_from_pptx(file_name)
        elif file_name.endswith('.pdf'):
            content = extract_text_from_pdf(file_name)
        else:
            content = f"Unknown file format: {file_name}"
        
        all_content.append(content)
        all_content.append("\n\n")
    
    # Write all content to a single file
    output_file = "course_content.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("\n".join(all_content))
    
    print(f"Content extracted successfully to {output_file}")

if __name__ == "__main__":
    main()