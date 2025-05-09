import os
import re
import sys
import json
import logging
from typing import Tuple, Set, List, Dict
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

input_path = sys.argv[1]
output_path = sys.argv[2]
json_output_path = "data/mapped-content.json"  # <-- Add this line

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

CONFIG = {
    'placeholder_prefix': '{{',
    'placeholder_suffix': '}}',
    'default_placeholder_patterns': [
        r'click\s+to\s+add',
        r'insert\s+your\s+text',
        r'type\s+here',
        r'your\s+text\s+here'
    ],
    'position_thresholds': {
        'header_top': 0.2,
        'bullet_indicator': ['•', '-', '*']
    },
    'type_mapping': {
        'title': 'TITLE',
        'subtitle': 'SUBTITLE',
        'header': 'HEADER',
        'bullet': 'BULLET_POINTS',
        'content': 'CONTENT',
        'image': 'IMAGE'
    }
}

def is_default_placeholder(text: str) -> bool:
    return any(re.search(pattern, text, re.IGNORECASE)
               for pattern in CONFIG['default_placeholder_patterns'])

def set_text_preserve_formatting(text_frame, new_text):
    if text_frame.paragraphs:
        para = text_frame.paragraphs[0]
        if para.runs:
            first_run = para.runs[0]
            font = first_run.font
            para.clear()
            new_run = para.add_run()
            new_run.text = new_text
            new_run.font.bold = font.bold
            new_run.font.italic = font.italic
            new_run.font.size = font.size
            if font.color and font.color.rgb:
                new_run.font.color.rgb = font.color.rgb
            return
    text_frame.text = new_text

def is_image_placeholder(shape) -> bool:
    try:
        return (
            shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and
            shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
        )
    except AttributeError:
        return False

def detect_placeholder_type(shape, text: str, slide_height: int) -> str:
    text_lower = text.lower()
    shape_name = shape.name.lower()

    if any(char in text for char in CONFIG['position_thresholds']['bullet_indicator']):
        return CONFIG['type_mapping']['bullet']
    if 'title' in shape_name:
        return CONFIG['type_mapping']['title']
    if 'subtitle' in shape_name:
        return CONFIG['type_mapping']['subtitle']
    if shape.top < slide_height * CONFIG['position_thresholds']['header_top']:
        return CONFIG['type_mapping']['header']
    if len(text.split('\n')) > 1:
        return CONFIG['type_mapping']['bullet']
    if len(text.split('. ')) > 2:
        return CONFIG['type_mapping']['content']
    return CONFIG['type_mapping']['content']

def generate_placeholder_name(placeholder_type: str, slide_num: int, index: int, slide_title: str) -> str:
    safe_title = re.sub(r'\W+', '_', slide_title.upper()) if slide_title else f'SLIDE_{slide_num}'
    return f"{CONFIG['placeholder_prefix']}{placeholder_type}_{index}_{safe_title}{CONFIG['placeholder_suffix']}"

def process_text_shape(shape, slide_num: int, slide_height: int, index: int, slide_title: str) -> str:
    try:
        text = shape.text.strip()
        if not text or is_default_placeholder(text):
            logger.debug(f"Skipping default text: '{text}'")
            return ""
        placeholder_type = detect_placeholder_type(shape, text, slide_height)
        if not placeholder_type:
            return ""
        new_name = generate_placeholder_name(placeholder_type, slide_num, index, slide_title)
        set_text_preserve_formatting(shape.text_frame, new_name)
        logger.info(f"Renamed '{text}' to '{new_name}'")
        return new_name
    except Exception as e:
        logger.warning(f"Error processing text shape: {str(e)}")
        return ""

def process_image_placeholder(shape, slide_num: int, index: int, slide_title: str) -> str:
    try:
        placeholder_type = CONFIG['type_mapping']['image']
        new_name = generate_placeholder_name(placeholder_type, slide_num, index, slide_title)
        logger.info(f"Found image placeholder: {new_name}")
        return new_name
    except Exception as e:
        logger.warning(f"Error processing image: {str(e)}")
        return ""

def process_grouped_shapes(group_shape, slide_num: int, slide_height: int, start_index: int, slide_title: str) -> List[str]:
    placeholders = []
    index = start_index
    shapes_sorted = sorted(group_shape.shapes, key=lambda s: s.top)
    for shape in shapes_sorted:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            inner, index = process_grouped_shapes(shape, slide_num, slide_height, index, slide_title)
            placeholders.extend(inner)
        elif shape.has_text_frame:
            result = process_text_shape(shape, slide_num, slide_height, index, slide_title)
            if result:
                placeholders.append(result)
                index += 1
    return placeholders, index

def process_pptx(input_path: str, output_path: str) -> Tuple[bool, Set[str], Dict[str, List[str]]]:
    try:
        if not os.path.exists(input_path):
            logger.error(f"Input file not found: {input_path}")
            return False, set(), {}

        prs = Presentation(input_path)
        placeholders = set()
        mapped = {}

        for slide_idx, slide in enumerate(prs.slides, 1):
            logger.info(f"\nProcessing slide {slide_idx}")
            slide_height = prs.slide_height
            slide_title = ""
            index = 1

            for shape in slide.shapes:
                if shape.has_text_frame and 'title' in shape.name.lower():
                    slide_title = shape.text.strip()
                    break

            shapes_sorted = sorted(slide.shapes, key=lambda s: s.top)
            slide_placeholders = []

            for shape in shapes_sorted:
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    results, index = process_grouped_shapes(shape, slide_idx, slide_height, index, slide_title)
                    slide_placeholders.extend(results)
                    placeholders.update(results)
                elif shape.has_text_frame:
                    result = process_text_shape(shape, slide_idx, slide_height, index, slide_title)
                    if result:
                        slide_placeholders.append(result)
                        placeholders.add(result)
                        index += 1
                elif is_image_placeholder(shape):
                    result = process_image_placeholder(shape, slide_idx, index, slide_title)
                    if result:
                        slide_placeholders.append(result)
                        placeholders.add(result)
                        index += 1

            if slide_placeholders:
                mapped[f"slide_{slide_idx}"] = slide_placeholders

        prs.save(output_path)
        logger.info(f"\nTotal placeholders processed: {len(placeholders)}")
        return True, placeholders, mapped

    except Exception as e:
        logger.error(f"\nProcessing failed: {str(e)}")
        return False, set(), {}

def write_json_mapping(mapping: Dict[str, List[str]], json_path: str):
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, indent=2, ensure_ascii=False)
    logger.info(f"\nExported placeholder mapping to {json_path}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python preprocess.py <input.pptx> <output.pptx>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    success, placeholders, mapping = process_pptx(input_file, output_file)

    if success:
        print("\nSuccessfully processed placeholders:")
        for ph in sorted(placeholders):
            print(f" - {ph}")

        # Export mapping to JSON
        write_json_mapping(mapping, json_output_path)

        sys.exit(0)
    else:
        print("\nProcessing completed with errors")
        sys.exit(1)

if __name__ == "__main__":
    main()
