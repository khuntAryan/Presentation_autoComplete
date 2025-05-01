from pptx import Presentation
import os
import re
import sys
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def detect_and_rename_placeholders(input_file, output_file):
    """Process a PowerPoint file to add {{PLACEHOLDER}} syntax to placeholder text"""
    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        return False
    
    logger.info(f"Processing file: {input_file}")
    prs = Presentation(input_file)
    
    # Track changes made
    changes_made = 0
    
    # Process each slide
    for slide_num, slide in enumerate(prs.slides, 1):
        logger.info(f"Processing slide {slide_num}")
        
        # Detect based on default PowerPoint placeholder text patterns
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text.strip()
                
                # First identify if this is already in {{PLACEHOLDER}} format
                if re.match(r'{{.*?}}', text):
                    logger.info(f"  Found existing placeholder: {text}")
                    continue
                
                # Check for common PowerPoint placeholder patterns
                placeholder_type = None
                
                # Placeholder titles are often in ALL CAPS or have "title" in their name
                if shape.name.lower().startswith("title") or "(TITLE)" in shape.text.upper():
                    placeholder_type = "TITLE"
                elif shape.name.lower().startswith("subtitle") or "(SUBTITLE)" in shape.text.upper():
                    placeholder_type = "SUBTITLE"
                elif "click to add text" in text.lower():
                    # Use shape position to determine type
                    if shape.top < 150:  # Near top of slide
                        placeholder_type = "HEADER"
                    else:
                        placeholder_type = f"CONTENT_{slide_num}_{len(text) % 10}"
                # Detect bullet list placeholders
                elif text.startswith("•") or text.startswith("*"):
                    placeholder_type = f"BULLETS_{slide_num}"
                # Default placeholder name based on location/slide number
                else:
                    if not text or "click to" in text.lower():
                        # Determine type by shape position on slide
                        if shape.top < 150:  # Near top of slide
                            placeholder_type = "HEADER"
                        elif shape.left < 150:  # Left side
                            placeholder_type = f"LEFT_CONTENT_{slide_num}"
                        elif shape.left > 400:  # Right side
                            placeholder_type = f"RIGHT_CONTENT_{slide_num}"
                        else:
                            placeholder_type = f"CONTENT_{slide_num}"
                
                # Set the text to the {{PLACEHOLDER}} format if type was determined
                if placeholder_type:
                    # Replace the text
                    new_text = f"{{{{{placeholder_type}}}}}"
                    shape.text_frame.text = new_text
                    changes_made += 1
                    logger.info(f"  Changed text to: {new_text}")
    
    # Save the processed file
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    prs.save(output_file)
    logger.info(f"Processing complete. Made {changes_made} changes.")
    logger.info(f"Saved to: {output_file}")
    return True

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python preprocess.py <input_pptx> <output_pptx>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    success = detect_and_rename_placeholders(input_file, output_file)
    if success:
        print(f"Successfully processed {input_file} to {output_file}")
    else:
        print("Processing failed")
        sys.exit(1)
