import re
import base64
import io
import xml.etree.ElementTree as ET
import logging
from typing import Dict, List, Tuple, Optional
from docx2python import docx2python
from PIL import Image
from xml.dom import minidom

# ============ LOGGING SETUP ============
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ============ REGEX PATTERNS ============
RE_OPTION = re.compile(r'^[A-E][\.\)]')
RE_IMAGE = re.compile(r'----(image\d+\.\w+)----')
RE_CLEAN_HTML = re.compile(r'<[^<]+?>')
RE_SANITIZE_FILENAME = re.compile(r'[<>:"/\\|?*]')
RE_INVALID_CHARS = re.compile(r'[<>&"\']')

# ============ CONSTANTS ============
MAX_IMAGE_WIDTH = 600
MIN_IMAGE_WIDTH = 100
IMAGE_QUALITY = 75
MAX_CATEGORY_LENGTH = 100
MAX_SNIPPET_LENGTH = 50
MAX_OPTIONS = 5
MIN_OPTIONS = 2

class ConversionError(Exception):
    """Custom exception untuk conversion errors"""
    pass

class ConversionStats:
    """Track conversion statistics"""
    def __init__(self):
        self.total_questions = 0
        self.total_images = 0
        self.warnings = []
        self.errors = []
    
    def add_warning(self, message: str):
        self.warnings.append(message)
        logger.warning(message)
    
    def add_error(self, message: str):
        self.errors.append(message)
        logger.error(message)
    
    def to_dict(self):
        return {
            'total_questions': self.total_questions,
            'total_images': self.total_images,
            'warnings': self.warnings,
            'errors': self.errors
        }

def cdata(text: str) -> str:
    """Wrapper untuk CDATA section"""
    if not text:
        return ""
    return f"<![CDATA[{text}]]>"

def optimize_image(
    image_bytes: bytes, 
    max_width: int = MAX_IMAGE_WIDTH,
    stats: Optional[ConversionStats] = None
) -> Tuple[bytes, str]:
    """
    Optimasi gambar dengan error handling yang comprehensive.
    
    Args:
        image_bytes: Raw image data
        max_width: Maximum width untuk resize
        stats: ConversionStats object untuk tracking
    
    Returns:
        Tuple(optimized_image_bytes, format_extension)
    """
    try:
        if not image_bytes:
            raise ValueError("Image bytes is empty")
        
        # Validate max_width
        max_width = max(MIN_IMAGE_WIDTH, min(max_width, 2000))
        
        # Open image dari bytes
        img = Image.open(io.BytesIO(image_bytes))
        
        # Handle berbagai image mode
        if img.mode in ("RGBA", "LA"):
            # Convert transparent ke white background
            background = Image.new("RGB", img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
            img = background
        elif img.mode == "P":
            img = img.convert("RGB")
        elif img.mode in ("L", "1"):
            # Grayscale dan binary
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")
        
        # Resize jika perlu
        if img.size[0] > max_width:
            aspect_ratio = img.size[1] / img.size[0]
            new_height = int(max_width * aspect_ratio)
            new_height = max(1, new_height)  # Ensure height > 0
            
            img = img.resize(
                (max_width, new_height),
                Image.Resampling.LANCZOS
            )
        
        # Save ke buffer dengan compression
        buffer = io.BytesIO()
        img.save(
            buffer,
            format="JPEG",
            quality=IMAGE_QUALITY,
            optimize=True
        )
        
        result_bytes = buffer.getvalue()
        
        if not result_bytes:
            raise ValueError("Failed to encode image to JPEG")
        
        logger.info(f"Image optimized: {len(image_bytes)} → {len(result_bytes)} bytes")
        return result_bytes, "jpg"
    
    except Exception as e:
        error_msg = f"Failed to optimize image: {str(e)}"
        logger.error(error_msg)
        
        if stats:
            stats.add_error(error_msg)
        
        # Return original bytes dengan format "png" sebagai fallback
        return image_bytes, "png"

def sanitize_filename(filename: str) -> str:
    """Sanitize filename untuk keamanan"""
    if not filename:
        return "quiz"
    
    # Remove dangerous characters
    safe_name = RE_SANITIZE_FILENAME.sub('_', filename)
    
    # Remove duplicate underscores
    safe_name = re.sub(r'_+', '_', safe_name)
    
    # Limit length
    safe_name = safe_name[:MAX_CATEGORY_LENGTH]
    
    # Remove trailing underscores
    safe_name = safe_name.rstrip('_')
    
    return safe_name or "quiz"

def sanitize_text(text: str, max_length: Optional[int] = None) -> str:
    """Sanitize text content untuk XML"""
    if not text:
        return ""
    
    text = text.strip()
    
    # Remove invalid XML characters tapi keep HTML tags
    # Keep hanya: alphanumeric, spaces, HTML tags, common punctuation
    if max_length:
        text = text[:max_length]
    
    return text

def validate_answer_format(answer_str: str) -> str:
    """
    Validasi dan normalize format jawaban.
    
    Input: "ANS: A, B", "A,B", "AB", etc
    Output: "AB" (uppercase letters A-E only)
    """
    if not answer_str:
        return ""
    
    # Remove "ANS:" prefix jika ada
    answer_str = answer_str.replace("ANS:", "").replace("ans:", "").strip()
    
    # Extract hanya huruf A-E
    valid_answers = ''.join(
        c.upper() for c in answer_str 
        if c.upper() in 'ABCDE'
    )
    
    return valid_answers

def parse_docx_to_questions(
    content,
    images_dict: Dict,
    stats: ConversionStats
) -> List[Dict]:
    """
    Parse DOCX content ke struktur questions.
    
    Returns:
        List of question dictionaries
    """
    questions = []
    current_q = None
    
    try:
        for sheet in content:
            for table in sheet:
                for row in table:
                    for cell in row:
                        for paragraph in cell:
                            text = sanitize_text(paragraph)
                            
                            if not text:
                                continue
                            
                            # Deteksi soal baru
                            if not RE_OPTION.match(text) and not text.startswith('ANS:'):
                                if current_q:
                                    # Validasi soal sebelum append
                                    if validate_question(current_q, stats):
                                        questions.append(current_q)
                                
                                q_type = 'essay' if 'ESSAY' in text.upper() else 'multichoice'
                                current_q = {
                                    'text': text,
                                    'options': [],
                                    'answer': '',
                                    'images': [],
                                    'type': q_type
                                }
                            
                            # Deteksi opsi jawaban
                            elif RE_OPTION.match(text) and current_q:
                                option_text = RE_OPTION.sub('', text).strip()
                                if option_text:  # Only add non-empty options
                                    current_q['options'].append(option_text)
                            
                            # Deteksi jawaban
                            elif text.startswith('ANS:') and current_q:
                                current_q['answer'] = validate_answer_format(text)
                            
                            # Deteksi gambar
                            if '----image' in text and current_q:
                                img_refs = RE_IMAGE.findall(text)
                                for ref in img_refs:
                                    if ref in images_dict:
                                        current_q['images'].append((ref, images_dict[ref]))
                                        stats.total_images += 1
                                    else:
                                        warning = f"Image reference '{ref}' not found in document"
                                        stats.add_warning(warning)
        
        # Append question terakhir
        if current_q and validate_question(current_q, stats):
            questions.append(current_q)
        
        stats.total_questions = len(questions)
        logger.info(f"Parsed {len(questions)} questions successfully")
        
    except Exception as e:
        error_msg = f"Error during question parsing: {str(e)}"
        stats.add_error(error_msg)
        raise ConversionError(error_msg)
    
    return questions

def validate_question(question: Dict, stats: ConversionStats) -> bool:
    """
    Validasi struktur question.
    
    Returns:
        True jika question valid, False jika ada issue
    """
    if not question.get('text'):
        stats.add_warning("Question has empty text")
        return False
    
    if question['type'] == 'multichoice':
        # Multichoice harus punya minimal 2 opsi
        if len(question['options']) < MIN_OPTIONS:
            warning = f"Multichoice question has only {len(question['options'])} options (min {MIN_OPTIONS})"
            stats.add_warning(warning)
            return False
        
        # Multichoice tidak boleh lebih dari 5 opsi
        if len(question['options']) > MAX_OPTIONS:
            warning = f"Multichoice question has {len(question['options'])} options (max {MAX_OPTIONS}), truncating"
            stats.add_warning(warning)
            question['options'] = question['options'][:MAX_OPTIONS]
        
        # Harus punya jawaban
        if not question['answer']:
            stats.add_warning("Multichoice question has no answer")
            return False
    
    return True

def process_docx_to_moodle(
    docx_file_bytes: bytes,
    filename: str = "QUIZ"
) -> Dict:
    """
    Proses DOCX file menjadi Moodle XML format.
    
    Args:
        docx_file_bytes: Raw DOCX file bytes
        filename: Original filename (untuk category name)
    
    Returns:
        Dictionary dengan keys:
            - 'xml': XML string result
            - 'stats': ConversionStats dictionary
            - 'success': Boolean status
    
    Raises:
        ConversionError: Jika terjadi error critical
    """
    stats = ConversionStats()
    
    try:
        # ============ INPUT VALIDATION ============
        if not docx_file_bytes:
            raise ConversionError("DOCX file bytes is empty")
        
        if len(docx_file_bytes) < 100:
            raise ConversionError("DOCX file is too small (possibly corrupted)")
        
        # Validate DOCX magic bytes (PK = ZIP format)
        if not docx_file_bytes.startswith(b'PK'):
            raise ConversionError("File is not a valid DOCX (invalid magic bytes)")
        
        # ============ EXTRACT CATEGORY NAME ============
        category_name = sanitize_filename(filename.rsplit('.', 1)[0])
        logger.info(f"Category name: {category_name}")
        
        # ============ PARSE DOCX ============
        try:
            with docx2python(docx_file_bytes, html=True) as doc_content:
                content = doc_content.body
                images_dict = doc_content.images
                
                if not content:
                    raise ConversionError("DOCX file has no content")
                
                logger.info(f"Found {len(images_dict)} images in document")
        
        except Exception as e:
            raise ConversionError(f"Failed to parse DOCX file: {str(e)}")
        
        # ============ CREATE ROOT XML ============
        quiz = ET.Element('quiz')
        
        # Add category
        cat_q = ET.SubElement(quiz, 'question', type='category')
        cat_node = ET.SubElement(cat_q, 'category')
        ET.SubElement(cat_node, 'text').text = f"$course$/{category_name}"
        
        # ============ PARSE QUESTIONS ============
        questions = parse_docx_to_questions(content, images_dict, stats)
        
        if not questions:
            raise ConversionError("No valid questions found in DOCX")
        
        # ============ BUILD XML STRUCTURE ============
        q_counter = 1
        
        for q in questions:
            try:
                q_node = ET.SubElement(quiz, 'question', type=q['type'])
                
                q_id = f"q{str(q_counter).zfill(3)}"
                pure_text = RE_CLEAN_HTML.sub('', q['text'])
                clean_snippet = pure_text[:MAX_SNIPPET_LENGTH].strip()
                
                # Build question name
                name_node = ET.SubElement(q_node, 'name')
                ET.SubElement(name_node, 'text').text = cdata(
                    f"{category_name} {q_id} {clean_snippet}"
                )
                
                # Build question text
                qtext_tag = ET.SubElement(q_node, 'questiontext', format='html')
                html_body = f"<p dir='auto'>{q['text']}</p>"
                
                # Process images
                for img_name, img_bytes in q['images']:
                    try:
                        opt_bytes, ext = optimize_image(img_bytes, MAX_IMAGE_WIDTH, stats)
                        new_img_name = f"{q_id}_{img_name.split('.')[0]}.{ext}"
                        img_html = (
                            f'<br /><img src="@@PLUGINFILE@@/{new_img_name}" '
                            f'alt="question image" border="0" /><br />'
                        )
                        html_body = html_body.replace(f'----{img_name}----', img_html)
                        
                        file_node = ET.SubElement(
                            qtext_tag, 'file',
                            name=new_img_name,
                            encoding="base64"
                        )
                        file_node.text = base64.b64encode(opt_bytes).decode()
                    
                    except Exception as e:
                        warning = f"Failed to process image {img_name}: {str(e)}"
                        stats.add_warning(warning)
                
                ET.SubElement(qtext_tag, 'text').text = cdata(html_body)
                
                # Default settings
                ET.SubElement(q_node, 'defaultgrade').text = "1.0"
                
                # Type-specific settings
                if q['type'] == 'multichoice':
                    ET.SubElement(q_node, 'penalty').text = "0.33"
                    ET.SubElement(q_node, 'hidden').text = "0"
                    ET.SubElement(q_node, 'single').text = "true"
                    ET.SubElement(q_node, 'shuffleanswers').text = "true"
                    ET.SubElement(q_node, 'answernumbering').text = "abc"
                    
                    letters = ['A', 'B', 'C', 'D', 'E']
                    for i, opt in enumerate(q['options']):
                        if i >= len(letters):
                            break
                        
                        is_correct = "100.0" if letters[i] in q['answer'] else "0.0"
                        ans_node = ET.SubElement(
                            q_node, 'answer',
                            fraction=is_correct,
                            format="html"
                        )
                        ET.SubElement(ans_node, 'text').text = cdata(opt)
                
                elif q['type'] == 'essay':
                    ET.SubElement(q_node, 'penalty').text = "0.1"
                    ET.SubElement(q_node, 'responseformat').text = "editor"
                    ET.SubElement(q_node, 'responserequired').text = "1"
                    ET.SubElement(q_node, 'responsefieldlines').text = "15"
                    ET.SubElement(q_node, 'attachments').text = "0"
                
                q_counter += 1
            
            except Exception as e:
                error_msg = f"Error processing question {q_counter}: {str(e)}"
                stats.add_error(error_msg)
                logger.error(error_msg)
                continue
        
        # ============ FINALIZE XML ============
        try:
            raw_xml = ET.tostring(quiz, encoding='utf-8')
            pretty_xml = minidom.parseString(raw_xml).toprettyxml(indent="  ")
            
            # Validate generated XML
            ET.fromstring(pretty_xml.encode('utf-8'))
            
            header_comment = (
                '<?xml version="1.0" encoding="UTF-8"?>\n'
                '<!-- Generated by Docx2Moodle Converter -->\n'
                '<!-- Questions: {} | Images: {} -->\n'
                '\n'.format(stats.total_questions, stats.total_images)
            )
            
            final_output = header_comment + pretty_xml.replace('<?xml version="1.0" ?>', '').strip()
            
            logger.info("XML generation successful")
            
            return {
                'xml': final_output,
                'stats': stats.to_dict(),
                'success': True
            }
        
        except ET.ParseError as e:
            error_msg = f"Generated invalid XML: {str(e)}"
            stats.add_error(error_msg)
            raise ConversionError(error_msg)
    
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"Unexpected error during conversion: {str(e)}"
        stats.add_error(error_msg)
        logger.error(error_msg, exc_info=True)
        raise ConversionError(error_msg)
