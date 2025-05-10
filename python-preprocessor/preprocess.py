import os
import re
import sys
import logging
import json
from typing import Tuple, Set, List, Dict
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

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
        'bullet': 'BULLET',
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
    return CONFIG['type_mapping']['content']

def generate_placeholder_name(placeholder_type: str, slide_num: int, index: int, word_count: int) -> str:
    """Generate placeholder name with position index and word count"""
    return f"{CONFIG['placeholder_prefix']}{placeholder_type}_{index}_SLIDE_{slide_num}_W{word_count}{CONFIG['placeholder_suffix']}"

def process_slide_shapes(slide, slide_num: int, slide_height: int) -> Tuple[List[str], int]:
    """Process all shapes in a slide with proper index tracking"""
    placeholders = []
    index = 1  # Reset index for each slide
    shapes_sorted = sorted(slide.shapes, key=lambda s: (s.top, s.left))

    for shape in shapes_sorted:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            group_placeholders, index = process_grouped_shapes(shape, slide_num, slide_height, index)
            placeholders.extend(group_placeholders)
        elif shape.has_text_frame:
            text = shape.text.strip()
            if not text or is_default_placeholder(text):
                continue
                
            placeholder_type = detect_placeholder_type(shape, text, slide_height)
            word_count = len(re.findall(r'\b[\w-]+\b', text))
            
            new_name = generate_placeholder_name(
                placeholder_type=placeholder_type,
                slide_num=slide_num,
                index=index,
                word_count=word_count
            )
            
            set_text_preserve_formatting(shape.text_frame, new_name)
            placeholders.append(new_name)
            index += 1
            logger.info(f"Renamed '{text}' to '{new_name}'")
            
        elif is_image_placeholder(shape):
            new_name = generate_placeholder_name(
                placeholder_type=CONFIG['type_mapping']['image'],
                slide_num=slide_num,
                index=index,
                word_count=0
            )
            placeholders.append(new_name)
            index += 1
            logger.info(f"Found image placeholder: {new_name}")

    return placeholders, index

def process_grouped_shapes(group_shape, slide_num: int, slide_height: int, start_index: int) -> Tuple[List[str], int]:
    """Process grouped shapes recursively"""
    placeholders = []
    index = start_index
    shapes_sorted = sorted(group_shape.shapes, key=lambda s: (s.top, s.left))
    
    for shape in shapes_sorted:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            inner, index = process_grouped_shapes(shape, slide_num, slide_height, index)
            placeholders.extend(inner)
        elif shape.has_text_frame:
            text = shape.text.strip()
            if not text or is_default_placeholder(text):
                continue
                
            placeholder_type = detect_placeholder_type(shape, text, slide_height)
            word_count = len(re.findall(r'\b[\w-]+\b', text))
            
            new_name = generate_placeholder_name(
                placeholder_type=placeholder_type,
                slide_num=slide_num,
                index=index,
                word_count=word_count
            )
            
            set_text_preserve_formatting(shape.text_frame, new_name)
            placeholders.append(new_name)
            index += 1
            logger.info(f"Renamed '{text}' to '{new_name}'")

    return placeholders, index

def process_pptx(input_path: str, output_path: str) -> Tuple[bool, Set[str], Dict[str, List[str]]]:
    """Main processing with sequential index tracking"""
    try:
        prs = Presentation(input_path)
        placeholders = set()
        mapped = {}

        for slide_idx, slide in enumerate(prs.slides, 1):
            logger.info(f"\nProcessing slide {slide_idx}")
            slide_placeholders, _ = process_slide_shapes(slide, slide_idx, prs.slide_height)
            placeholders.update(slide_placeholders)
            mapped[f"slide_{slide_idx}"] = slide_placeholders

        prs.save(output_path)
        logger.info(f"\nTotal placeholders processed: {len(placeholders)}")
        return True, placeholders, mapped
    except Exception as e:
        logger.error(f"\nProcessing failed: {str(e)}")
        return False, set(), {}

def write_json_mapping(mapping: Dict[str, List[str]], json_path: str):
    """Save mapping to JSON file"""
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, indent=2, ensure_ascii=False)
    logger.info(f"\nExported placeholder mapping to {json_path}")

def generate_ai_prompt(mapping: Dict[str, List[str]], output_path: str, topic: str = "[INSERT YOUR TOPIC HERE]"):
    """Generate AI prompt template for plain text output (not JSON)"""
    prompt_lines = [
        "Please generate content for a PowerPoint presentation based on the following structure.",
        f"Presentation Topic: {topic}",
        "",
        "Instructions:",
        "1. Write the output in plain text, slide by slide, as shown in the example.",
        "2. Use the placeholder types and recommended word counts as a guide.",
        "3. It's OK to use fewer words than the word count, but do not go over.",
        "4. For bullets, use a new line for each point.",
        "5. Do NOT return JSON or any code block, only text.",
        "",
        "Slide Structure and Placeholders:"
    ]

    for slide_key in sorted(mapping.keys(), key=lambda x: int(x.split('_')[1])):
        slide_num = slide_key.split('_')[1]
        placeholders = mapping[slide_key]
        prompt_lines.append(f"\nSlide {slide_num}:")
        for ph in placeholders:
            match = re.match(r'\{\{(\w+)_(\d+)_SLIDE_(\d+)_W(\d+)\}\}', ph)
            if match:
                ph_type, ph_index, slide_num, word_count = match.groups()
                prompt_lines.append(
                    f"  - {ph_type.title()} (max {word_count} words)"
                )

    prompt_lines.extend([
        "",
        "Example Output:",
        "Slide 1: Title Slide",
        "Title: Life on Land – Sustainable Development Goal 15",
        "Subtitle: Protecting, Restoring, and Promoting Terrestrial Ecosystems",
        "Presented by: [Your Name/Class/Institution]",
        "",
        "Slide 2: Objectives",
        "Understand the importance of Life on Land (SDG 15).",
        "Identify threats to terrestrial ecosystems.",
        "Explore actions and solutions to preserve biodiversity and land.",
        "Encourage individual and community involvement.",
        "",
        "Slide 3: What is Life on Land?",
        "SDG 15 focuses on conserving forests, combating desertification, halting land degradation, and protecting biodiversity.",
        "Healthy land ecosystems are vital for food, water, air, and climate stability.",
        "Supports life for more than 80% of terrestrial species.",
        "",
        "Slide 4: Importance of Forests and Land Ecosystems",
        "Forests cover 31% of Earth’s land area.",
        "Provide oxygen, regulate climate, prevent erosion, and support millions of species.",
        "Forests are also crucial for human livelihoods (1.6 billion people rely on them).",
        "",
        "Slide 5: Threats to Life on Land",
        "Deforestation (for agriculture, logging).",
        "Desertification (due to overgrazing, climate change).",
        "Pollution (soil contamination, plastics).",
        "Urbanization (loss of green spaces and ecosystems).",
        "Biodiversity loss (extinction due to habitat destruction).",
        "",
        "Slide 6: Impact of Biodiversity Loss",
        "Reduces ecosystem resilience.",
        "Affects food chains and natural processes (like pollination).",
        "Increases the risk of pandemics from disrupted wildlife.",
        "Threatens human survival and well-being.",
        "",
        "Slide 7: Global Efforts",
        "UN Convention on Biological Diversity.",
        "Bonn Challenge: Restore 350 million hectares by 2030.",
        "Protected Areas: 15% of land is under protection globally.",
        "Countries adopting policies for sustainable forestry and farming.",
        "",
        "Slide 8: What Can We Do?",
        "Reduce paper and wood consumption.",
        "Support reforestation and conservation programs.",
        "Avoid products from illegal logging or endangered species.",
        "Raise awareness and participate in local clean-up/planting events.",
        "",
        "Slide 9: Success Stories",
        "Costa Rica: Doubled forest cover through eco-policies and tourism.",
        "India’s Chipko Movement: Community-based forest protection.",
        "Great Green Wall in Africa: Reclaiming land across 20+ countries.",
        "",
        "Slide 10: Conclusion & Call to Action",
        "Healthy land is the foundation for a healthy planet.",
        "Every action counts-preserve, protect, and restore.",
        "Let’s build a future where both people and nature thrive.",
        "Quote: “The Earth is what we all have in common.” – Wendell Berry",
        "",
        "Remember: Do NOT use JSON. Write in text, slide by slide, as above."
    ])

    prompt = "\n".join(prompt_lines)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(prompt)
    logger.info(f"\nGenerated AI prompt template at: {output_path}")

def get_placeholder_description(ph_type: str) -> str:
    descriptions = {
        'TITLE': "concise and attention-grabbing title text",
        'SUBTITLE': "brief supporting statement or tagline",
        'HEADER': "section header text",
        'BULLET': "bullet points list using '•'",
        'CONTENT': "detailed paragraph content",
        'IMAGE': "image description/caption"
    }
    return descriptions.get(ph_type, "general content")

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
        
        write_json_mapping(mapping, "data/mapped-content.json")
        generate_ai_prompt(mapping, "data/ai-prompt-template.txt")
        
        print("\nNext steps:")
        print("1. Edit 'data/ai-prompt-template.txt' with your topic")
        print("2. Use the prompt with your preferred AI service")
        print("3. Save the AI response to 'data/user-content.txt' (plain text)")
        print("4. Run the presentation generator")
        sys.exit(0)
    else:
        print("\nProcessing completed with errors")
        sys.exit(1)

if __name__ == "__main__":
    main()
