from pptx import Presentation
import os
import re
import sys
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def detect_and_rename_placeholders(input_file, output_file):
    """Process a PowerPoint file to add {{PLACEHOLDER}} syntax to placeholder text dynamically"""
    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        return False
    
    logger.info(f"Processing file: {input_file}")
    prs = Presentation(input_file)
    
    changes_made = 0
    
    for slide_num, slide in enumerate(prs.slides, 1):
        logger.info(f"Processing slide {slide_num}")
        
        # Initialize counters for each type in this slide
        counters = {
            "TITLE": 0,
            "SUBTITLE": 0,
            "HEADER": 0,
            "CONTENT": 0,
            "BULLETS": 0,
            "LEFT_CONTENT": 0,
            "RIGHT_CONTENT": 0,
        }
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text.strip()
                
                if re.match(r'{{.*?}}', text):
                    logger.info(f"  Found existing placeholder: {text}")
                    continue
                
                placeholder_type = None
                
                if shape.name.lower().startswith("title") or "(TITLE)" in shape.text.upper():
                    placeholder_type = "TITLE"
                elif shape.name.lower().startswith("subtitle") or "(SUBTITLE)" in shape.text.upper():
                    placeholder_type = "SUBTITLE"
                elif "click to add text" in text.lower():
                    if shape.top < 150:
                        placeholder_type = "HEADER"
                    else:
                        placeholder_type = "CONTENT"
                elif text.startswith("•") or text.startswith("*"):
                    placeholder_type = "BULLETS"
                else:
                    if not text or "click to" in text.lower():
                        if shape.top < 150:
                            placeholder_type = "HEADER"
                        elif shape.left < 150:
                            placeholder_type = "LEFT_CONTENT"
                        elif shape.left > 400:
                            placeholder_type = "RIGHT_CONTENT"
                        else:
                            placeholder_type = "CONTENT"
                
                if placeholder_type:
                    # Increment counter
                    counters[placeholder_type] += 1
                    
                    if placeholder_type in ["TITLE", "SUBTITLE", "HEADER"]:
                        placeholder_name = f"{{{{{placeholder_type}_SLIDE_{slide_num}}}}}"
                    else:
                        placeholder_name = f"{{{{{placeholder_type}_SLIDE_{slide_num}_{counters[placeholder_type]}}}}}"
                    
                    shape.text_frame.text = placeholder_name
                    changes_made += 1
                    logger.info(f"  Changed text to: {placeholder_name}")
    
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
