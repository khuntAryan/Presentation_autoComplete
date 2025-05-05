from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
import os
import re
import sys
import logging
from typing import Tuple, Set

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

DEFAULT_PLACEHOLDER_PATTERNS = [
    r'click\s+to\s+add',
    r'insert\s+your\s+text',
    r'type\s+here',
    r'your\s+text\s+here'
]

def set_text_preserve_formatting(text_frame, new_text):
    """Replace text while preserving first run's formatting"""
    if not text_frame.paragraphs:
        text_frame.text = new_text
        return

    para = text_frame.paragraphs[0]
    if not para.runs:
        text_frame.text = new_text
        return

    # Keep first run's formatting
    first_run = para.runs[0]
    first_run.text = new_text

    # Remove additional runs if any
    for run in para.runs[1:]:
        p = run._element.getparent()
        p.remove(run._element)

def is_image_placeholder(shape) -> bool:
    """Check if shape is a PowerPoint image placeholder"""
    try:
        return (
            shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and
            shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        )
    except AttributeError:
        return False

def process_image_placeholder(shape, slide_num: int) -> str:
    """Process image placeholders"""
    try:
        placeholder_type = "IMAGE"
        new_name = f"{{{{{placeholder_type}_SLIDE_{slide_num}}}}}"
        logger.info(f"Found image placeholder: {new_name}")
        return new_name
    except Exception as e:
        logger.warning(f"Error processing image: {str(e)}")
        return ""

def process_text_shape(shape, slide_num: int) -> str:
    """Process text shapes with actual content (preserves formatting)"""
    try:
        text = shape.text.strip()

        # Skip empty/default text boxes
        if not text or is_default_placeholder(text):
            logger.debug(f"Skipping default text: '{text}'")
            return ""

        # Detect placeholder type
        placeholder_type = detect_placeholder_type(shape, text)
        if not placeholder_type:
            return ""

        # Generate standardized name
        new_name = generate_placeholder_name(placeholder_type, slide_num)

        # Preserve formatting while replacing text
        set_text_preserve_formatting(shape.text_frame, new_name)

        logger.info(f"Renamed '{text}' to '{new_name}'")
        return new_name

    except Exception as e:
        logger.warning(f"Error processing text shape: {str(e)}")
        return ""

def is_default_placeholder(text: str) -> bool:
    """Check for default placeholder patterns"""
    return any(re.search(pattern, text, re.IGNORECASE)
               for pattern in DEFAULT_PLACEHOLDER_PATTERNS)

def detect_placeholder_type(shape, text: str) -> str:
    """Detect placeholder type using multiple strategies"""
    # Check shape name
    shape_name = shape.name.lower()
    if 'title' in shape_name:
        return "TITLE"
    if 'subtitle' in shape_name:
        return "SUBTITLE"

    # Check text patterns
    if re.search(r'\b(title|heading)\b', text, re.IGNORECASE):
        return "HEADER"
    if re.search(r'\b(content|body)\b', text, re.IGNORECASE):
        return "CONTENT"

    # Position-based detection
    if shape.top < 100:
        return "HEADER"

    return "CUSTOM"

def generate_placeholder_name(placeholder_type: str, slide_num: int) -> str:
    """Generate standardized placeholder name"""
    return f"{{{{{placeholder_type}_SLIDE_{slide_num}}}}}"

def process_pptx(input_path: str, output_path: str) -> Tuple[bool, Set[str]]:
    """Process PPTX file to standardize placeholders"""
    try:
        if not os.path.exists(input_path):
            logger.error(f"Input file not found: {input_path}")
            return False, set()

        prs = Presentation(input_path)
        placeholders = set()
        changes_made = 0

        for slide_idx, slide in enumerate(prs.slides, 1):
            logger.info(f"\nProcessing slide {slide_idx}")

            for shape in slide.shapes:
                logger.debug(f"Checking shape: {shape.name} ({shape.shape_type})")

                # Process text placeholders
                if shape.has_text_frame:
                    result = process_text_shape(shape, slide_idx)
                    if result:
                        placeholders.add(result)
                        changes_made += 1
                # Process image placeholders
                elif is_image_placeholder(shape):
                    result = process_image_placeholder(shape, slide_idx)
                    if result:
                        placeholders.add(result)
                        changes_made += 1

        prs.save(output_path)
        logger.info(f"\nTotal placeholders processed: {changes_made}")
        return True, placeholders

    except Exception as e:
        logger.error(f"\nProcessing failed: {str(e)}")
        return False, set()

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python preprocess.py <input.pptx> <output.pptx>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    success, placeholders = process_pptx(input_file, output_file)

    if success:
        print("\nSuccessfully processed placeholders:")
        for ph in sorted(placeholders):
            print(f" - {ph}")
        sys.exit(0)
    else:
        print("\nProcessing completed with errors")
        sys.exit(1)
