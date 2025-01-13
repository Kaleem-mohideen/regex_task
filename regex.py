from docx import Document
import re

def check_word_shading_in_italic_text(doc):
    highlighted_numbers = []
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Check if the run is italic
            if run.font.italic:
                print(f"Inspecting italic text: '{run.text}'")
                
                # Split the run text into words
                words = run.text.split()
                
                # Check for shading on the run level
                shading_element = run._element.xpath(".//w:shd")
                if shading_element:
                    shading_val = shading_element[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
                    if shading_val:
                        # Print each word from the italicized run that has shading
                        for word in words:
                            print(f"Word '{word}' has shading with color: #{shading_val}")
                            if shading_val.upper() == "FF9900":  # Orange peel color
                                # highlighted_numbers.append(word)
                                highlighted_numbers.extend(re.findall(r'\b\d+\b', word))
                                print(f"Word '{word}' has orange peel shading.")
                else:
                    # If the run has no shading
                    for word in words:
                        print(f"Word '{word}' in italic text has no shading.")
    return highlighted_numbers
# Example usage
file_path = "Django Evaluation - Final.docx"
doc = Document(file_path)
highlight_bacgrnd = check_word_shading_in_italic_text(doc)
print("Extracted numbers with orange peel shading:", highlight_bacgrnd)


# # Call the function and print results
# numbers = extract_numbers_with_highlighting(file_path)
# print("Extracted numbers with highlighting or shading:", numbers)
