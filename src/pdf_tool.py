"""PDF file manipulation tool for Gemini Office Agent.

This module provides the PdfTool class for reading, creating, and modifying
PDF files using PyPDF2 (read/merge/split/rotate/metadata) and reportlab (create).
"""

import io
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from PyPDF2.errors import PdfReadError
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.colors import HexColor, black, gray
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, PageBreak
)
from reportlab.pdfgen import canvas

from src.exceptions import CorruptedFileError
from src.logging_config import get_logger, log_file_access_error

logger = get_logger(__name__)


class PdfTool:
    """Tool for manipulating PDF files.
    
    Provides methods for reading, creating, merging, splitting,
    and modifying PDF files using PyPDF2 and reportlab.
    """

    def read_pdf(self, file_path: str) -> Dict[str, Any]:
        """Extrai texto de todas as páginas do PDF.
        
        Args:
            file_path: Path to the PDF file
            
        Returns:
            Dictionary containing:
                - pages: List of text content per page
                - num_pages: Total number of pages
                - metadata: PDF metadata dict
                
        Raises:
            FileNotFoundError: If the file does not exist
            CorruptedFileError: If the file is corrupted
        """
        logger.info(f"Reading PDF file: {file_path}")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        try:
            reader = PdfReader(file_path)
            
            pages_text = []
            for page in reader.pages:
                text = page.extract_text() or ""
                pages_text.append(text)
            
            metadata = {}
            if reader.metadata:
                metadata = {
                    "title": reader.metadata.title,
                    "author": reader.metadata.author,
                    "subject": reader.metadata.subject,
                    "creator": reader.metadata.creator,
                }
            
            result = {
                "pages": pages_text,
                "num_pages": len(reader.pages),
                "metadata": metadata
            }
            
            logger.info(f"Successfully read PDF with {len(reader.pages)} pages")
            return result
            
        except PdfReadError as e:
            logger.error(f"Corrupted PDF file: {file_path} - {str(e)}")
            raise CorruptedFileError(f"PDF file is corrupted or invalid: {file_path}") from e
        except Exception as e:
            logger.error(f"Error reading PDF file: {file_path} - {str(e)}")
            raise CorruptedFileError(f"PDF file is corrupted or invalid: {file_path} - {str(e)}") from e

    def create_pdf(self, file_path: str, elements: List[Dict[str, Any]],
                   page_size: str = "A4") -> None:
        """Cria PDF estruturado com múltiplos tipos de elementos.
        
        Args:
            file_path: Path where the PDF should be created
            elements: List of dicts, each with 'type' and type-specific fields:
                - {"type": "title", "text": "..."}
                - {"type": "heading", "text": "...", "level": 1}
                - {"type": "paragraph", "text": "...", "alignment": "left"}
                - {"type": "table", "headers": [...], "rows": [[...]]}
                - {"type": "list", "items": [...], "ordered": false}
                - {"type": "spacer", "height": 20}
                - {"type": "page_break"}
            page_size: "A4" or "letter" (default: "A4")
            
        Raises:
            IOError: If the file cannot be created
        """
        logger.info(f"Creating PDF file: {file_path} ({len(elements)} elements)")
        
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)
        
        try:
            size = A4 if page_size.upper() == "A4" else letter
            doc = SimpleDocTemplate(
                file_path,
                pagesize=size,
                rightMargin=2 * cm,
                leftMargin=2 * cm,
                topMargin=2 * cm,
                bottomMargin=2 * cm
            )
            
            styles = getSampleStyleSheet()
            
            # Estilos customizados
            styles.add(ParagraphStyle(
                name='CustomTitle',
                parent=styles['Title'],
                fontSize=24,
                spaceAfter=30,
            ))
            styles.add(ParagraphStyle(
                name='CustomH1',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=12,
            ))
            styles.add(ParagraphStyle(
                name='CustomH2',
                parent=styles['Heading2'],
                fontSize=14,
                spaceAfter=10,
            ))
            styles.add(ParagraphStyle(
                name='CustomH3',
                parent=styles['Heading3'],
                fontSize=12,
                spaceAfter=8,
            ))
            
            flowables = []
            
            for elem in elements:
                elem_type = elem.get('type', 'paragraph')
                
                if elem_type == 'title':
                    flowables.append(Paragraph(elem.get('text', ''), styles['CustomTitle']))
                
                elif elem_type == 'heading':
                    level = elem.get('level', 1)
                    style_name = f'CustomH{min(level, 3)}'
                    flowables.append(Paragraph(elem.get('text', ''), styles[style_name]))
                
                elif elem_type == 'paragraph':
                    text = elem.get('text', '')
                    alignment = elem.get('alignment', 'left')
                    align_map = {
                        'left': TA_LEFT, 'center': TA_CENTER,
                        'right': TA_RIGHT, 'justify': TA_JUSTIFY,
                    }
                    
                    para_style = ParagraphStyle(
                        name=f'para_{id(elem)}',
                        parent=styles['Normal'],
                        alignment=align_map.get(alignment, TA_LEFT),
                        fontSize=elem.get('font_size', 11),
                        spaceAfter=6,
                    )
                    
                    # Suportar negrito e itálico via tags
                    if elem.get('bold'):
                        text = f'<b>{text}</b>'
                    if elem.get('italic'):
                        text = f'<i>{text}</i>'
                    
                    flowables.append(Paragraph(text, para_style))
                
                elif elem_type == 'table':
                    headers = elem.get('headers', [])
                    rows = elem.get('rows', [])
                    
                    table_data = [headers] + rows
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#4472C4')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#FFFFFF')),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 11),
                        ('FONTSIZE', (0, 1), (-1, -1), 10),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('GRID', (0, 0), (-1, -1), 0.5, gray),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#F2F2F2'), HexColor('#FFFFFF')]),
                        ('TOPPADDING', (0, 0), (-1, -1), 6),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ]))
                    flowables.append(Spacer(1, 6))
                    flowables.append(t)
                    flowables.append(Spacer(1, 6))
                
                elif elem_type == 'list':
                    items = elem.get('items', [])
                    ordered = elem.get('ordered', False)
                    
                    bullet_type = '1' if ordered else 'bullet'
                    list_items = [
                        ListItem(Paragraph(item, styles['Normal']))
                        for item in items
                    ]
                    flowables.append(ListFlowable(list_items, bulletType=bullet_type))
                
                elif elem_type == 'spacer':
                    height = elem.get('height', 20)
                    flowables.append(Spacer(1, height))
                
                elif elem_type == 'page_break':
                    flowables.append(PageBreak())
            
            doc.build(flowables)
            logger.info(f"Successfully created PDF with {len(elements)} elements")
            
        except Exception as e:
            logger.error(f"Error creating PDF: {file_path} - {str(e)}")
            raise IOError(f"Failed to create PDF: {file_path} - {str(e)}") from e

    def merge_pdfs(self, file_paths: List[str], output_path: str) -> None:
        """Junta múltiplos PDFs em um único arquivo.
        
        Args:
            file_paths: List of PDF file paths to merge (in order)
            output_path: Path for the merged output file
            
        Raises:
            FileNotFoundError: If any input file does not exist
            IOError: If merge fails
        """
        logger.info(f"Merging {len(file_paths)} PDFs into {output_path}")
        
        for fp in file_paths:
            if not Path(fp).exists():
                log_file_access_error(logger, fp, "arquivo não encontrado")
                raise FileNotFoundError(f"PDF file not found: {fp}")
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        try:
            merger = PdfMerger()
            for fp in file_paths:
                merger.append(fp)
            merger.write(output_path)
            merger.close()
            
            logger.info(f"Successfully merged {len(file_paths)} PDFs")
            
        except PdfReadError as e:
            raise CorruptedFileError(f"One of the PDF files is corrupted: {e}") from e
        except Exception as e:
            logger.error(f"Error merging PDFs: {str(e)}")
            raise IOError(f"Failed to merge PDFs: {str(e)}") from e

    def split_pdf(self, file_path: str, output_path: str,
                  start_page: int, end_page: int) -> None:
        """Divide PDF extraindo um range de páginas.
        
        Args:
            file_path: Path to the source PDF
            output_path: Path for the output PDF
            start_page: First page to extract (1-based)
            end_page: Last page to extract (1-based, inclusive)
            
        Raises:
            FileNotFoundError: If the file does not exist
            ValueError: If page range is invalid
        """
        logger.info(f"Splitting PDF {file_path}: pages {start_page}-{end_page}")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        if start_page < 1:
            raise ValueError(f"start_page must be >= 1, got {start_page}")
        if end_page < start_page:
            raise ValueError(f"end_page ({end_page}) must be >= start_page ({start_page})")
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        try:
            reader = PdfReader(file_path)
            
            if end_page > len(reader.pages):
                raise ValueError(
                    f"end_page ({end_page}) exceeds total pages ({len(reader.pages)})"
                )
            
            writer = PdfWriter()
            # Convert to 0-based index
            for i in range(start_page - 1, end_page):
                writer.add_page(reader.pages[i])
            
            with open(output_path, 'wb') as f:
                writer.write(f)
            
            logger.info(f"Successfully split PDF: {end_page - start_page + 1} pages extracted")
            
        except PdfReadError as e:
            raise CorruptedFileError(f"PDF file is corrupted: {file_path}") from e
        except ValueError:
            raise
        except Exception as e:
            logger.error(f"Error splitting PDF: {file_path} - {str(e)}")
            raise IOError(f"Failed to split PDF: {file_path} - {str(e)}") from e

    def add_text_overlay(self, file_path: str, text: str,
                         x: float = 200, y: float = 400,
                         font_size: int = 40, color: str = "CCCCCC",
                         opacity: float = 0.3,
                         pages: Optional[List[int]] = None) -> None:
        """Adiciona texto/marca d'água sobre páginas existentes.
        
        Args:
            file_path: Path to the PDF file
            text: Text to overlay
            x: X position in points (default 200)
            y: Y position in points (default 400)
            font_size: Font size (default 40)
            color: Hex color without # (default "CCCCCC" - light gray)
            opacity: Opacity 0.0-1.0 (default 0.3)
            pages: List of page numbers (1-based) to apply to, or None for all pages
            
        Raises:
            FileNotFoundError: If the file does not exist
        """
        logger.info(f"Adding text overlay '{text[:30]}' to {file_path}")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()
            
            # Parse color
            r = int(color[0:2], 16) / 255.0
            g = int(color[2:4], 16) / 255.0
            b = int(color[4:6], 16) / 255.0
            
            for i, page in enumerate(reader.pages):
                page_num = i + 1  # 1-based
                
                # Verificar se deve aplicar nesta página
                if pages is not None and page_num not in pages:
                    writer.add_page(page)
                    continue
                
                # Criar overlay com canvas
                packet = io.BytesIO()
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)
                c = canvas.Canvas(packet, pagesize=(page_width, page_height))
                
                c.saveState()
                c.setFillColorRGB(r, g, b)
                c.setFont("Helvetica", font_size)
                
                # Rotacionar 45 graus para watermark diagonal
                c.translate(x, y)
                c.rotate(45)
                c.setFillAlpha(opacity)
                c.drawString(0, 0, text)
                
                c.restoreState()
                c.save()
                
                packet.seek(0)
                overlay_reader = PdfReader(packet)
                overlay_page = overlay_reader.pages[0]
                
                page.merge_page(overlay_page)
                writer.add_page(page)
            
            with open(file_path, 'wb') as f:
                writer.write(f)
            
            logger.info(f"Successfully added text overlay to PDF")
            
        except PdfReadError as e:
            raise CorruptedFileError(f"PDF file is corrupted: {file_path}") from e
        except Exception as e:
            logger.error(f"Error adding text overlay: {file_path} - {str(e)}")
            raise IOError(f"Failed to add text overlay: {file_path} - {str(e)}") from e

    def get_info(self, file_path: str) -> Dict[str, Any]:
        """Obtém informações e metadata do PDF.
        
        Args:
            file_path: Path to the PDF file
            
        Returns:
            Dictionary with PDF metadata and info
            
        Raises:
            FileNotFoundError: If the file does not exist
        """
        logger.info(f"Getting info for PDF: {file_path}")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        try:
            reader = PdfReader(file_path)
            
            info = {
                "num_pages": len(reader.pages),
                "file_size_bytes": path.stat().st_size,
                "file_size_kb": round(path.stat().st_size / 1024, 1),
                "encrypted": reader.is_encrypted,
            }
            
            if reader.metadata:
                info.update({
                    "title": reader.metadata.title,
                    "author": reader.metadata.author,
                    "subject": reader.metadata.subject,
                    "creator": reader.metadata.creator,
                    "producer": reader.metadata.producer,
                })
            
            # Tamanho da primeira página
            if reader.pages:
                page = reader.pages[0]
                info["page_width"] = float(page.mediabox.width)
                info["page_height"] = float(page.mediabox.height)
            
            logger.info(f"Successfully retrieved PDF info: {info['num_pages']} pages")
            return info
            
        except PdfReadError as e:
            raise CorruptedFileError(f"PDF file is corrupted: {file_path}") from e
        except Exception as e:
            logger.error(f"Error getting PDF info: {file_path} - {str(e)}")
            raise IOError(f"Failed to get PDF info: {file_path} - {str(e)}") from e

    def rotate_pages(self, file_path: str, rotation: int,
                     pages: Optional[List[int]] = None) -> None:
        """Rotaciona páginas do PDF.
        
        Args:
            file_path: Path to the PDF file
            rotation: Rotation angle (90, 180, or 270)
            pages: List of page numbers (1-based) to rotate, or None for all
            
        Raises:
            FileNotFoundError: If the file does not exist
            ValueError: If rotation is invalid
        """
        logger.info(f"Rotating pages in {file_path} by {rotation} degrees")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        if rotation not in (90, 180, 270):
            raise ValueError(f"Rotation must be 90, 180, or 270, got {rotation}")
        
        try:
            reader = PdfReader(file_path)
            writer = PdfWriter()
            
            for i, page in enumerate(reader.pages):
                page_num = i + 1  # 1-based
                if pages is None or page_num in pages:
                    page.rotate(rotation)
                writer.add_page(page)
            
            with open(file_path, 'wb') as f:
                writer.write(f)
            
            pages_desc = f"pages {pages}" if pages else "all pages"
            logger.info(f"Successfully rotated {pages_desc} by {rotation} degrees")
            
        except PdfReadError as e:
            raise CorruptedFileError(f"PDF file is corrupted: {file_path}") from e
        except ValueError:
            raise
        except Exception as e:
            logger.error(f"Error rotating pages: {file_path} - {str(e)}")
            raise IOError(f"Failed to rotate pages: {file_path} - {str(e)}") from e

    def extract_tables(self, file_path: str) -> List[List[List[str]]]:
        """Extrai tabelas do PDF baseado em padrões de texto.
        
        Usa heurística simples: detecta linhas com separadores consistentes
        (tabs, múltiplos espaços) como tabelas.
        
        Args:
            file_path: Path to the PDF file
            
        Returns:
            List of tables, each table is a list of rows, each row is a list of cells
            
        Raises:
            FileNotFoundError: If the file does not exist
        """
        logger.info(f"Extracting tables from PDF: {file_path}")
        
        path = Path(file_path)
        if not path.exists():
            log_file_access_error(logger, file_path, "arquivo não encontrado")
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        try:
            reader = PdfReader(file_path)
            tables = []
            
            for page in reader.pages:
                text = page.extract_text() or ""
                lines = text.split('\n')
                
                current_table = []
                for line in lines:
                    line = line.strip()
                    if not line:
                        if current_table and len(current_table) >= 2:
                            tables.append(current_table)
                        current_table = []
                        continue
                    
                    # Detectar separadores: tab ou múltiplos espaços
                    if '\t' in line:
                        cells = [c.strip() for c in line.split('\t') if c.strip()]
                    elif '  ' in line:
                        import re
                        cells = [c.strip() for c in re.split(r'\s{2,}', line) if c.strip()]
                    else:
                        if current_table and len(current_table) >= 2:
                            tables.append(current_table)
                        current_table = []
                        continue
                    
                    if len(cells) >= 2:
                        current_table.append(cells)
                    else:
                        if current_table and len(current_table) >= 2:
                            tables.append(current_table)
                        current_table = []
                
                if current_table and len(current_table) >= 2:
                    tables.append(current_table)
            
            logger.info(f"Extracted {len(tables)} table(s) from PDF")
            return tables
            
        except PdfReadError as e:
            raise CorruptedFileError(f"PDF file is corrupted: {file_path}") from e
        except Exception as e:
            logger.error(f"Error extracting tables from PDF: {file_path} - {str(e)}")
            raise IOError(f"Failed to extract tables: {file_path} - {str(e)}") from e
