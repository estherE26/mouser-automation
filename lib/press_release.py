"""
Press Release Processing Module (Serverless-adapted)
Processes press release folders and generates HTML files.
"""

import os
import re
import tempfile
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from ftplib import FTP

from docx import Document
from docx.oxml.ns import qn
from PIL import Image, ImageFilter


# ============================================================================
# CONFIGURATION
# ============================================================================

DEFAULT_CONFIG = {
    "ftp": {
        "host": os.environ.get("FTP_HOST", "3.143.159.140"),
        "port": 21,
        "username": os.environ.get("FTP_USER"),
        "password": os.environ.get("FTP_PASS"),
        "base_remote_path": "/Mouser/{month_folder}/{folder_name}/"
    },
    "image_settings": {
        "output_width": 336,
        "jpeg_quality": 95
    },
    "urls": {
        "base_url": "https://pr.ezwire.com/Mouser/{month_folder}/{folder_name}/"
    },
    "contacts": {
        "marketing": {
            "name": "Kevin Hess",
            "title": "Senior Vice President of Marketing",
            "company": "Mouser Electronics",
            "phone": "(817) 804-3833",
            "email": "Kevin.Hess@mouser.com"
        },
        "press": {
            "name": "Kelly DeGarmo",
            "title": "Manager, Corporate Communications and Media Relations",
            "company": "Mouser Electronics",
            "phone": "(817) 804-7764",
            "email": "Kelly.DeGarmo@mouser.com"
        }
    }
}


# ============================================================================
# DOCUMENT PARSING
# ============================================================================

def clean_url(url: str) -> str:
    """Remove tracking parameters from URLs for display."""
    if not url:
        return url
    if '?' in url:
        return url.split('?')[0]
    return url


def extract_hyperlinks(paragraph) -> Dict[str, str]:
    """Extract hyperlinks from a paragraph, mapping text to URLs."""
    hyperlinks = {}
    p_xml = paragraph._element

    for hyperlink in p_xml.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
        r_id = hyperlink.get(qn('r:id'))
        if r_id:
            try:
                rel = paragraph.part.rels.get(r_id)
                if rel and rel.target_ref:
                    link_text = ''.join(
                        node.text for node in hyperlink.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        if node.text
                    )
                    if link_text:
                        hyperlinks[link_text] = rel.target_ref
            except Exception:
                pass

    return hyperlinks


def get_paragraph_html(paragraph, include_links: bool = True) -> str:
    """Convert a paragraph to HTML, preserving formatting and links."""
    hyperlinks = extract_hyperlinks(paragraph) if include_links else {}
    text = paragraph.text

    if hyperlinks and include_links:
        link_positions = []
        for link_text, url in hyperlinks.items():
            pos = text.find(link_text)
            if pos >= 0:
                link_positions.append((pos, len(link_text), link_text, url))

        link_positions.sort(key=lambda x: (-x[0], -x[1]))
        linked_ranges = []
        result_text = text

        for pos, length, link_text, url in link_positions:
            overlaps = False
            for start, end in linked_ranges:
                if not (pos + length <= start or pos >= end):
                    overlaps = True
                    break

            if not overlaps:
                clean = clean_url(url)
                link_html = f'<a href="{clean}" target="_blank">{link_text}</a>'
                result_text = result_text[:pos] + link_html + result_text[pos + length:]
                linked_ranges.append((pos, pos + length))

        return result_text

    return text


def parse_docx(file_path: str) -> Dict:
    """Parse a Word document and extract structured content."""
    doc = Document(file_path)

    result = {
        'headline': '',
        'subheadline': '',
        'date': '',
        'body_paragraphs': [],
        'about_sections': {},
        'product_link': '',
        'meta_description': '',
        'meta_keywords': ''
    }

    current_section = 'body'
    current_about_title = None

    for para in doc.paragraphs:
        text = para.text.strip()
        style_name = para.style.name if para.style else 'Normal'

        if not text:
            continue

        if text in ['– 30 –', '- 30 -']:
            continue

        if style_name == 'Title' or (style_name == 'Title2' and not result['headline']):
            if 'New Product Announcement' in text:
                continue
            result['headline'] = text
            continue

        if not result['headline'] and style_name in ['Title', 'Heading 1'] and len(text) > 20:
            result['headline'] = text
            continue

        if result['headline'] and not result['subheadline'] and not result['body_paragraphs']:
            if style_name == 'Subtitle':
                result['subheadline'] = text
                continue
            has_italic = any(run.italic for run in para.runs if run.text.strip())
            if has_italic and len(text) < 200 and text != result['headline']:
                if not re.match(r'^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d', text):
                    result['subheadline'] = text
                    continue

        if style_name == 'Heading 1' or text.startswith('About '):
            if text.startswith('About '):
                current_section = 'about'
                current_about_title = text
                result['about_sections'][current_about_title] = []
                continue

        if 'Trademarks' in text and (style_name == 'Heading 1' or text.startswith('Trademarks')):
            current_section = 'trademarks'
            continue

        if current_section == 'trademarks':
            continue

        if current_section == 'about' and current_about_title:
            para_html = get_paragraph_html(para)
            result['about_sections'][current_about_title].append(para_html)
        elif current_section == 'body':
            para_html = get_paragraph_html(para)

            date_match = re.match(r'^((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})', para_html)
            if date_match and not result['date']:
                result['date'] = date_match.group(1)

            if not result['product_link']:
                links = extract_hyperlinks(para)
                for url in links.values():
                    if 'mouser.com' in url and '/new/' in url:
                        result['product_link'] = clean_url(url)
                        break

            result['body_paragraphs'].append(para_html)

    if result['body_paragraphs']:
        first_para = result['body_paragraphs'][0]
        clean_text = re.sub(r'<[^>]+>', '', first_para)
        result['meta_description'] = clean_text[:160].rsplit(' ', 1)[0] + '...' if len(clean_text) > 160 else clean_text

    if result['headline']:
        words = re.findall(r'\b[A-Z][a-zA-Z0-9-]+\b', result['headline'])
        result['meta_keywords'] = ', '.join(words[:10]).lower()

    return result


# ============================================================================
# IMAGE PROCESSING
# ============================================================================

def convert_png_to_jpg(
    png_path: str,
    output_path: str,
    target_width: int = 336,
    jpeg_quality: int = 95
) -> Dict:
    """Convert PNG to JPG with resizing."""
    result = {
        'success': False,
        'output_path': output_path,
        'dimensions': None,
        'error': None
    }

    try:
        with Image.open(png_path) as img:
            ratio = target_width / img.width
            new_height = int(img.height * ratio)

            if img.mode in ('RGBA', 'LA', 'P'):
                rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    rgb_img.paste(img, mask=img.split()[-1])
                else:
                    rgb_img.paste(img)
                img = rgb_img
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            resized = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
            sharpened = resized.filter(ImageFilter.UnsharpMask(radius=1.5, percent=120, threshold=2))
            sharpened.save(output_path, 'JPEG', quality=jpeg_quality, optimize=True)

            result['success'] = True
            result['dimensions'] = (target_width, new_height)

    except Exception as e:
        result['error'] = str(e)

    return result


# ============================================================================
# HTML GENERATION
# ============================================================================

def generate_month_folder(date_str: str = None) -> str:
    """Generate month folder name like '2026-01 - Mouser'."""
    if date_str:
        try:
            date_obj = datetime.strptime(date_str, "%B %d, %Y")
            return f"{date_obj.year}-{date_obj.month:02d} - Mouser"
        except ValueError:
            pass
    now = datetime.now()
    return f"{now.year}-{now.month:02d} - Mouser"


def format_body_paragraphs(paragraphs: list, for_email: bool = False) -> str:
    """Format body paragraphs as HTML."""
    if for_email:
        return '\n'.join(
            f'<p style="font:14px Helvetica, Arial, sans-serif; line-height: 20px;">{p}</p>'
            for p in paragraphs
        )
    return '\n'.join(f'                    <p>{p}</p>' for p in paragraphs)


def format_about_sections(about_sections: Dict, for_email: bool = False) -> str:
    """Format 'About' sections as HTML."""
    if not about_sections:
        return ''

    formatted = []
    for title, paragraphs in about_sections.items():
        if for_email:
            formatted.append(f'<h2 style="font:16px Helvetica, Arial, sans-serif; line-height: 20px; font-weight: bold;"><u>{title}</u></h2>')
            for para in paragraphs:
                formatted.append(f'<p style="font:14px Helvetica, Arial, sans-serif; line-height: 20px;">{para}</p>')
        else:
            formatted.append(f'                    <h1><u>{title}</u></h1>')
            for para in paragraphs:
                formatted.append(f'                    <p>{para}</p>')

    return '\n'.join(formatted)


def generate_press_release_html(
    template: str,
    content: Dict,
    folder_name: str,
    jpg_filename: str,
    png_filename: str,
    pdf_filename: str,
    image_dimensions: tuple,
    config: Dict,
    product_link: str = None,
    image_url: str = None
) -> str:
    """Generate the main press release HTML."""
    month_folder = generate_month_folder(content.get('date'))
    base_url = config['urls']['base_url'].format(
        month_folder=month_folder.replace(' ', '%20'),
        folder_name=folder_name.replace(' ', '%20')
    )

    jpg_url = image_url if image_url else f"{base_url}{jpg_filename}"
    png_url = f"{base_url}{png_filename}"
    pdf_url = f"{base_url}{pdf_filename}"

    if not product_link:
        product_link = content.get('product_link', '#')

    body_html = format_body_paragraphs(content['body_paragraphs'], for_email=False)
    about_html = format_about_sections(content['about_sections'], for_email=False)

    marketing = config['contacts']['marketing']
    press = config['contacts']['press']

    replacements = {
        '{{title}}': folder_name,
        '{{meta_description}}': content.get('meta_description', ''),
        '{{meta_keywords}}': content.get('meta_keywords', ''),
        '{{headline}}': content.get('headline', ''),
        '{{subheadline}}': content.get('subheadline', ''),
        '{{jpg_url}}': jpg_url,
        '{{png_url}}': png_url,
        '{{pdf_url}}': pdf_url,
        '{{product_link}}': product_link,
        '{{image_alt}}': content.get('headline', 'Product Image'),
        '{{image_width}}': str(image_dimensions[0]),
        '{{image_height}}': str(image_dimensions[1]),
        '{{body_paragraphs}}': body_html,
        '{{about_sections}}': about_html,
        '{{contact_marketing_name}}': marketing['name'],
        '{{contact_marketing_company}}': marketing['company'],
        '{{contact_marketing_title}}': marketing['title'],
        '{{contact_marketing_phone}}': marketing['phone'],
        '{{contact_marketing_email}}': marketing['email'],
        '{{contact_press_name}}': press['name'],
        '{{contact_press_company}}': press['company'],
        '{{contact_press_title}}': press['title'],
        '{{contact_press_phone}}': press['phone'],
        '{{contact_press_email}}': press['email'],
    }

    result = template
    for placeholder, value in replacements.items():
        result = result.replace(placeholder, value)

    return result


def generate_email_html(
    template: str,
    content: Dict,
    folder_name: str,
    jpg_filename: str,
    png_filename: str,
    pdf_filename: str,
    image_dimensions: tuple,
    config: Dict,
    product_link: str = None,
    image_url: str = None,
    subject: str = None
) -> str:
    """Generate the email version HTML."""
    month_folder = generate_month_folder(content.get('date'))
    base_url = config['urls']['base_url'].format(
        month_folder=month_folder.replace(' ', '%20'),
        folder_name=folder_name.replace(' ', '%20')
    )

    jpg_url = image_url if image_url else f"{base_url}{jpg_filename}"
    png_url = f"{base_url}{png_filename}"
    pdf_url = f"{base_url}{pdf_filename}"
    web_version_url = f"{base_url}{folder_name}.html"

    if not product_link:
        product_link = content.get('product_link', '#')

    email_subject = subject if subject else content.get('headline', '')

    paragraphs = content['body_paragraphs']
    first_paragraph = paragraphs[0] if paragraphs else ''
    remaining_paragraphs = format_body_paragraphs(paragraphs[1:], for_email=True) if len(paragraphs) > 1 else ''
    about_html = format_about_sections(content['about_sections'], for_email=True)

    marketing = config['contacts']['marketing']
    press = config['contacts']['press']

    replacements = {
        '{{title}}': folder_name,
        '{{subject}}': email_subject,
        '{{web_version_url}}': web_version_url,
        '{{headline}}': content.get('headline', ''),
        '{{subheadline}}': content.get('subheadline', ''),
        '{{jpg_url}}': jpg_url,
        '{{png_url}}': png_url,
        '{{pdf_url}}': pdf_url,
        '{{product_link}}': product_link,
        '{{image_alt}}': content.get('headline', 'Product Image'),
        '{{image_width}}': str(image_dimensions[0]),
        '{{image_height}}': str(image_dimensions[1]),
        '{{first_paragraph}}': first_paragraph,
        '{{remaining_paragraphs}}': remaining_paragraphs,
        '{{about_sections_email}}': about_html,
        '{{contact_marketing_name}}': marketing['name'],
        '{{contact_marketing_company}}': marketing['company'],
        '{{contact_marketing_title}}': marketing['title'],
        '{{contact_marketing_phone}}': marketing['phone'],
        '{{contact_marketing_email}}': marketing['email'],
        '{{contact_press_name}}': press['name'],
        '{{contact_press_company}}': press['company'],
        '{{contact_press_title}}': press['title'],
        '{{contact_press_phone}}': press['phone'],
        '{{contact_press_email}}': press['email'],
    }

    result = template
    for placeholder, value in replacements.items():
        result = result.replace(placeholder, value)

    return result


# ============================================================================
# FTP UPLOAD
# ============================================================================

def upload_to_ftp(files: List[Dict], config: Dict, folder_name: str, month_folder: str) -> Dict:
    """Upload files to FTP server."""
    ftp_config = config['ftp']
    remote_path = ftp_config['base_remote_path'].format(
        month_folder=month_folder,
        folder_name=folder_name
    )

    result = {'success': False, 'error': None, 'uploaded': 0}

    try:
        ftp = FTP()
        ftp.connect(ftp_config['host'], ftp_config['port'])
        ftp.login(ftp_config['username'], ftp_config['password'])

        # Create directory structure
        dirs = remote_path.strip('/').split('/')
        current = ''
        for d in dirs:
            current += '/' + d
            try:
                ftp.cwd(current)
            except:
                ftp.mkd(current)

        ftp.cwd('/')

        # Upload files
        for file_info in files:
            local_path = file_info['local_path']
            remote_filename = file_info.get('remote_filename', os.path.basename(local_path))
            full_remote = f"{remote_path}{remote_filename}"

            with open(local_path, 'rb') as f:
                ftp.storbinary(f'STOR {full_remote}', f)
            result['uploaded'] += 1

        ftp.quit()
        result['success'] = True

    except Exception as e:
        result['error'] = str(e)

    return result


# ============================================================================
# MAIN PROCESSING
# ============================================================================

def find_files(folder_path: str) -> Dict:
    """Find required files in the press release folder."""
    files = {
        'docx': None,
        'png': None,
        'pdf': None,
        'folder_name': os.path.basename(folder_path)
    }

    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        if not os.path.isfile(filepath):
            continue

        lower = filename.lower()

        if lower.endswith('.docx') and 'instruction' not in lower:
            if files['docx'] is None or 'publicrelations' in lower:
                files['docx'] = filepath

        elif lower.endswith('.png'):
            files['png'] = filepath

        elif lower.endswith('.pdf'):
            if 'instruction' not in lower and 'order' not in lower:
                if files['pdf'] is None or 'publicrelations' in lower:
                    files['pdf'] = filepath

    return files


def process_press_release(
    folder_path: str,
    press_release_template: str,
    email_template: str,
    image_url: str = None,
    subject: str = None,
    config: Dict = None
) -> Dict:
    """
    Process a press release folder and generate HTML files.

    Args:
        folder_path: Path to the press release folder
        press_release_template: HTML template content for press release
        email_template: HTML template content for email
        image_url: Optional tracking URL for embedded image
        subject: Optional email subject line
        config: Optional config dict (uses DEFAULT_CONFIG if not provided)

    Returns:
        Dictionary with processing results
    """
    config = config or DEFAULT_CONFIG

    result = {
        'success': False,
        'folder_name': None,
        'files_to_upload': [],
        'preview_urls': {},
        'errors': [],
        'month_folder': None
    }

    # Validate folder exists
    if not os.path.isdir(folder_path):
        result['errors'].append(f"Folder not found: {folder_path}")
        return result

    folder_name = os.path.basename(folder_path)
    result['folder_name'] = folder_name

    # Find required files
    files = find_files(folder_path)
    missing = []
    if not files['docx']:
        missing.append("Word document (.docx)")
    if not files['png']:
        missing.append("PNG image (.png)")
    if not files['pdf']:
        missing.append("PDF file (.pdf)")

    if missing:
        result['errors'].append(f"Missing files: {', '.join(missing)}")
        return result

    # Parse Word document
    try:
        content = parse_docx(files['docx'])
    except Exception as e:
        result['errors'].append(f"Failed to parse Word document: {e}")
        return result

    # Process image
    png_basename = os.path.splitext(os.path.basename(files['png']))[0]
    jpg_path = os.path.join(folder_path, f"{png_basename}.jpg")

    img_result = convert_png_to_jpg(
        files['png'],
        jpg_path,
        target_width=config['image_settings']['output_width'],
        jpeg_quality=config['image_settings']['jpeg_quality']
    )

    if not img_result['success']:
        result['errors'].append(f"Failed to process image: {img_result['error']}")
        return result

    image_dimensions = img_result['dimensions']

    # Generate HTML files
    jpg_filename = os.path.basename(jpg_path)
    png_filename = os.path.basename(files['png'])
    pdf_filename = os.path.basename(files['pdf'])

    html_filename = f"{folder_name}.html"
    email_filename = f"{folder_name}_email.html"

    product_link = content.get('product_link', '#')

    try:
        # Generate main press release HTML
        press_release_html = generate_press_release_html(
            press_release_template,
            content,
            folder_name,
            jpg_filename,
            png_filename,
            pdf_filename,
            image_dimensions,
            config,
            product_link=product_link,
            image_url=image_url
        )

        html_path = os.path.join(folder_path, html_filename)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(press_release_html)

        # Generate email HTML
        email_html = generate_email_html(
            email_template,
            content,
            folder_name,
            jpg_filename,
            png_filename,
            pdf_filename,
            image_dimensions,
            config,
            product_link=product_link,
            image_url=image_url,
            subject=subject
        )

        email_path = os.path.join(folder_path, email_filename)
        with open(email_path, 'w', encoding='utf-8') as f:
            f.write(email_html)

    except Exception as e:
        result['errors'].append(f"Failed to generate HTML: {e}")
        return result

    # Build file list for upload
    result['files_to_upload'] = [
        {'local_path': html_path, 'remote_filename': html_filename},
        {'local_path': email_path, 'remote_filename': email_filename},
        {'local_path': jpg_path, 'remote_filename': jpg_filename},
        {'local_path': files['png'], 'remote_filename': png_filename},
        {'local_path': files['pdf'], 'remote_filename': pdf_filename},
    ]

    result['month_folder'] = generate_month_folder(content.get('date'))

    # Generate preview URLs
    base_url = config['urls']['base_url'].format(
        month_folder=result['month_folder'].replace(' ', '%20'),
        folder_name=folder_name.replace(' ', '%20')
    )
    result['preview_urls'] = {
        'html': f"{base_url}{html_filename}",
        'email': f"{base_url}{email_filename}"
    }

    result['success'] = True
    return result
