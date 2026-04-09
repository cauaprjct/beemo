"""Conversor de HTML para elementos estruturados do Word.

Quando o modelo Gemini retorna HTML no campo 'content' ao invés de usar
o array 'elements', este módulo converte as tags HTML em elementos
compatíveis com o create_structured do WordTool.
"""

import re
from typing import Any, Dict, List, Optional
from html.parser import HTMLParser

from src.logging_config import get_logger

logger = get_logger(__name__)


def contains_html(text: str) -> bool:
    """Verifica se o texto contém tags HTML significativas.
    
    Retorna True se encontrar tags HTML como <h1>, <p>, <strong>, <table>, etc.
    Ignora tags soltas ou texto que não é realmente HTML.
    """
    if not text or not isinstance(text, str):
        return False
    
    # Padrão para detectar tags HTML comuns que o modelo costuma gerar
    html_pattern = re.compile(
        r'<\s*(?:h[1-6]|p|strong|em|b|i|u|ul|ol|li|table|tr|td|th|br|hr|div|span|a)\b[^>]*>',
        re.IGNORECASE
    )
    
    matches = html_pattern.findall(text)
    # Precisa de pelo menos 2 tags para considerar HTML (evita falsos positivos)
    return len(matches) >= 2


class _HTMLToElementsParser(HTMLParser):
    """Parser HTML que converte tags em elementos estruturados."""
    
    def __init__(self):
        super().__init__()
        self.elements: List[Dict[str, Any]] = []
        self._current_text = ""
        self._tag_stack: List[str] = []
        self._current_bold = False
        self._current_italic = False
        self._current_underline = False
        self._current_alignment = None
        self._list_type: Optional[str] = None  # 'ul' or 'ol'
        self._list_items: List[str] = []
        self._table_headers: List[str] = []
        self._table_rows: List[List[str]] = []
        self._current_row: List[str] = []
        self._in_header_row = False
        self._in_table = False
        self._in_list = False
        self._in_link = False
        self._link_url = ""
    
    def handle_starttag(self, tag: str, attrs: list):
        tag = tag.lower()
        attrs_dict = dict(attrs)
        self._tag_stack.append(tag)
        
        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self._flush_text()
            self._current_text = ""
        
        elif tag == 'p':
            self._flush_text()
            self._current_text = ""
            # Check for alignment in style
            style = attrs_dict.get('style', '')
            if 'text-align' in style:
                if 'center' in style:
                    self._current_alignment = 'center'
                elif 'right' in style:
                    self._current_alignment = 'right'
                elif 'justify' in style:
                    self._current_alignment = 'justify'
        
        elif tag in ('strong', 'b'):
            self._current_bold = True
        
        elif tag in ('em', 'i'):
            self._current_italic = True
        
        elif tag == 'u':
            self._current_underline = True
        
        elif tag in ('ul', 'ol'):
            self._flush_text()
            self._in_list = True
            self._list_type = tag
            self._list_items = []
        
        elif tag == 'li':
            self._current_text = ""
        
        elif tag == 'table':
            self._flush_text()
            self._in_table = True
            self._table_headers = []
            self._table_rows = []
            self._in_header_row = True
        
        elif tag == 'tr':
            self._current_row = []
        
        elif tag in ('th', 'td'):
            self._current_text = ""
            if tag == 'th':
                self._in_header_row = True
        
        elif tag == 'br':
            self._current_text += "\n"
        
        elif tag == 'a':
            self._in_link = True
            self._link_url = attrs_dict.get('href', '')
    
    def handle_endtag(self, tag: str):
        tag = tag.lower()
        
        if self._tag_stack and self._tag_stack[-1] == tag:
            self._tag_stack.pop()
        
        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            level = int(tag[1])
            text = self._current_text.strip()
            if text:
                self.elements.append({
                    "type": "heading",
                    "text": text,
                    "level": level - 1 if level == 1 else level  # h1 -> level 0 (Title)
                })
            self._current_text = ""
        
        elif tag == 'p':
            text = self._current_text.strip()
            if text:
                elem = {
                    "type": "paragraph",
                    "text": text,
                }
                if self._current_bold:
                    elem["bold"] = True
                if self._current_italic:
                    elem["italic"] = True
                if self._current_alignment:
                    elem["alignment"] = self._current_alignment
                self.elements.append(elem)
            self._current_text = ""
            self._current_bold = False
            self._current_italic = False
            self._current_alignment = None
        
        elif tag in ('strong', 'b'):
            self._current_bold = False
        
        elif tag in ('em', 'i'):
            self._current_italic = False
        
        elif tag == 'u':
            self._current_underline = False
        
        elif tag == 'li':
            text = self._current_text.strip()
            if text:
                self._list_items.append(text)
            self._current_text = ""
        
        elif tag in ('ul', 'ol'):
            if self._list_items:
                self.elements.append({
                    "type": "list",
                    "items": self._list_items,
                    "ordered": tag == 'ol'
                })
            self._list_items = []
            self._in_list = False
            self._list_type = None
        
        elif tag in ('th', 'td'):
            text = self._current_text.strip()
            self._current_row.append(text)
            self._current_text = ""
        
        elif tag == 'tr':
            if self._in_header_row and self._current_row:
                self._table_headers = self._current_row
                self._in_header_row = False
            elif self._current_row:
                self._table_rows.append(self._current_row)
            self._current_row = []
        
        elif tag == 'table':
            if self._table_headers or self._table_rows:
                # If no explicit headers, use first row
                if not self._table_headers and self._table_rows:
                    self._table_headers = self._table_rows.pop(0)
                self.elements.append({
                    "type": "table",
                    "headers": self._table_headers,
                    "rows": self._table_rows
                })
            self._in_table = False
            self._table_headers = []
            self._table_rows = []
        
        elif tag == 'a':
            self._in_link = False
            self._link_url = ""
    
    def handle_data(self, data: str):
        self._current_text += data
    
    def _flush_text(self):
        """Flush any pending text as a paragraph."""
        text = self._current_text.strip()
        if text and not self._in_list and not self._in_table:
            self.elements.append({
                "type": "paragraph",
                "text": text
            })
        self._current_text = ""


def html_to_elements(html_content: str) -> List[Dict[str, Any]]:
    """Converte HTML em lista de elementos estruturados.
    
    Args:
        html_content: String com conteúdo HTML
        
    Returns:
        Lista de dicts compatíveis com WordTool.create_structured()
    """
    logger.info("Convertendo conteúdo HTML para elementos estruturados")
    
    parser = _HTMLToElementsParser()
    
    try:
        parser.feed(html_content)
        # Flush remaining text
        parser._flush_text()
    except Exception as e:
        logger.warning(f"Erro ao parsear HTML: {e}. Usando fallback de texto simples.")
        # Fallback: strip tags and return as plain paragraphs
        return _fallback_strip_html(html_content)
    
    elements = parser.elements
    
    if not elements:
        # Parser não gerou nada, usar fallback
        return _fallback_strip_html(html_content)
    
    logger.info(f"HTML convertido em {len(elements)} elementos estruturados")
    return elements


def _fallback_strip_html(html_content: str) -> List[Dict[str, Any]]:
    """Remove tags HTML e retorna como parágrafos simples.
    
    Último recurso quando o parser HTML falha.
    """
    # Remove tags HTML
    clean = re.sub(r'<[^>]+>', '\n', html_content)
    # Limpa espaços extras
    clean = re.sub(r'\n\s*\n', '\n', clean).strip()
    
    elements = []
    for line in clean.split('\n'):
        line = line.strip()
        if line:
            elements.append({
                "type": "paragraph",
                "text": line
            })
    
    return elements if elements else [{"type": "paragraph", "text": ""}]
