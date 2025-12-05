"""
Style-related functions for Word Document Server.
"""
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE


def ensure_heading_style(doc):
    """
    Ensure Heading styles exist in the document.
    
    Args:
        doc: Document object
    """
    for i in range(1, 10):  # Create Heading 1 through Heading 9
        style_name = f'Heading {i}'
        try:
            # Try to access the style to see if it exists
            style = doc.styles[style_name]
        except KeyError:
            # Create the style if it doesn't exist
            try:
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                if i == 1:
                    style.font.size = Pt(16)
                    style.font.bold = True
                elif i == 2:
                    style.font.size = Pt(14)
                    style.font.bold = True
                else:
                    style.font.size = Pt(12)
                    style.font.bold = True
            except Exception:
                # If style creation fails, we'll just use default formatting
                pass


def ensure_table_style(doc):
    """
    Ensure Table Grid style exists in the document.
    
    Args:
        doc: Document object
    """
    try:
        # Try to access the style to see if it exists
        style = doc.styles['Table Grid']
    except KeyError:
        # If style doesn't exist, we'll handle it at usage time
        pass


def ensure_paragraph_style(doc):
    """
    Ensure Parágrafo (Body Text) style exists in the document.
    Creates it if it doesn't exist with proper formatting.
    
    Args:
        doc: Document object
        
    Returns:
        The style name to use
    """
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import logging
    
    # Log all available paragraph styles for debugging
    available_styles = [s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]
    logging.info(f"[Styles] Available paragraph styles: {available_styles}")
    
    # Try to find existing paragraph style
    # Order matters: try Portuguese names first, then English
    style_names_to_try = [
        'Parágrafo',           # Portuguese custom
        'Paragrafo',           # Portuguese without accent
        'Corpo de texto',      # Portuguese Body Text
        'Corpo do texto',      # Portuguese variant
        'Body Text',           # English
        'Body',                # Short English
        'Normal',              # Default fallback
    ]
    
    for style_name in style_names_to_try:
        try:
            style = doc.styles[style_name]
            logging.info(f"[Styles] Found paragraph style: '{style_name}'")
            return style_name
        except KeyError:
            continue
    
    # Create 'Parágrafo' style if none exists
    try:
        style = doc.styles.add_style('Parágrafo', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Cm(1.25)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.line_spacing = 1.5
        return 'Parágrafo'
    except Exception:
        # If creation fails, return None to use manual formatting
        return None


def ensure_list_paragraph_style(doc):
    """
    Ensure List Paragraph style exists in the document.
    This is the proper style for bullet and numbered lists in Word.
    
    Args:
        doc: Document object
        
    Returns:
        The style name to use for list items
    """
    from docx.shared import Pt, Cm
    import logging
    
    # Try to find existing list paragraph style (Portuguese and English variants)
    # Order matters: prefer the style that exists in the template
    style_names_to_try = [
        'List Paragraph',      # English (common in templates)
        'Parágrafo da Lista',  # Portuguese
        'Párrafo de lista',    # Spanish
    ]
    
    for style_name in style_names_to_try:
        try:
            style = doc.styles[style_name]
            logging.info(f"[Styles] Found existing list style: '{style_name}'")
            return style_name
        except KeyError:
            continue
    
    # Log all available styles for debugging
    logging.info(f"[Styles] Available styles in document: {[s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]}")
    
    # Create 'Parágrafo da Lista' style if none exists (Portuguese name to match Word PT-BR)
    try:
        style = doc.styles.add_style('Parágrafo da Lista', WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        style.paragraph_format.left_indent = Cm(1.27)  # Standard list indent
        style.paragraph_format.space_after = Pt(3)
        style.paragraph_format.space_before = Pt(0)
        logging.info(f"[Styles] Created new list style: 'Parágrafo da Lista'")
        return 'Parágrafo da Lista'
    except Exception as e:
        logging.error(f"[Styles] Failed to create list style: {e}")
        # If creation fails, return None to use manual formatting
        return None


def create_style(doc, style_name, style_type, base_style=None, font_properties=None, paragraph_properties=None):
    """
    Create a new style in the document.
    
    Args:
        doc: Document object
        style_name: Name for the new style
        style_type: Type of style (WD_STYLE_TYPE)
        base_style: Optional base style to inherit from
        font_properties: Dictionary of font properties (bold, italic, size, name, color)
        paragraph_properties: Dictionary of paragraph properties (alignment, spacing)
        
    Returns:
        The created style
    """
    from docx.shared import Pt
    
    try:
        # Check if style already exists
        style = doc.styles.get_by_id(style_name, WD_STYLE_TYPE.PARAGRAPH)
        return style
    except:
        # Create new style
        new_style = doc.styles.add_style(style_name, style_type)
        
        # Set base style if specified
        if base_style:
            new_style.base_style = doc.styles[base_style]
        
        # Set font properties
        if font_properties:
            font = new_style.font
            if 'bold' in font_properties:
                font.bold = font_properties['bold']
            if 'italic' in font_properties:
                font.italic = font_properties['italic']
            if 'size' in font_properties:
                font.size = Pt(font_properties['size'])
            if 'name' in font_properties:
                font.name = font_properties['name']
            if 'color' in font_properties:
                from docx.shared import RGBColor
                
                # Define common RGB colors
                color_map = {
                    'red': RGBColor(255, 0, 0),
                    'blue': RGBColor(0, 0, 255),
                    'green': RGBColor(0, 128, 0),
                    'yellow': RGBColor(255, 255, 0),
                    'black': RGBColor(0, 0, 0),
                    'gray': RGBColor(128, 128, 128),
                    'white': RGBColor(255, 255, 255),
                    'purple': RGBColor(128, 0, 128),
                    'orange': RGBColor(255, 165, 0)
                }
                
                color_value = font_properties['color']
                try:
                    # Handle string color names
                    if isinstance(color_value, str) and color_value.lower() in color_map:
                        font.color.rgb = color_map[color_value.lower()]
                    # Handle RGBColor objects
                    elif hasattr(color_value, 'rgb'):
                        font.color.rgb = color_value
                    # Try to parse as RGB string
                    elif isinstance(color_value, str):
                        font.color.rgb = RGBColor.from_string(color_value)
                    # Use directly if it's already an RGB value
                    else:
                        font.color.rgb = color_value
                except Exception as e:
                    # Fallback to black if all else fails
                    font.color.rgb = RGBColor(0, 0, 0)
        
        # Set paragraph properties
        if paragraph_properties:
            if 'alignment' in paragraph_properties:
                new_style.paragraph_format.alignment = paragraph_properties['alignment']
            if 'spacing' in paragraph_properties:
                new_style.paragraph_format.line_spacing = paragraph_properties['spacing']
        
        return new_style
