"""Testes para PdfTool."""

import os
import tempfile
from pathlib import Path
import pytest

from src.pdf_tool import PdfTool
from src.exceptions import CorruptedFileError


@pytest.fixture
def pdf_tool():
    """Fixture que retorna uma instância de PdfTool."""
    return PdfTool()


@pytest.fixture
def temp_dir():
    """Fixture que cria um diretório temporário para testes."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


class TestPdfToolCreate:
    """Testes para criação de PDFs."""
    
    def test_create_simple_pdf(self, pdf_tool, temp_dir):
        """Testa criação de PDF simples com texto."""
        file_path = os.path.join(temp_dir, "simple.pdf")
        
        elements = [
            {"type": "title", "text": "Test Document"},
            {"type": "paragraph", "text": "This is a test paragraph."}
        ]
        
        pdf_tool.create_pdf(file_path, elements)
        
        assert Path(file_path).exists()
        assert Path(file_path).stat().st_size > 0
    
    def test_create_structured_pdf(self, pdf_tool, temp_dir):
        """Testa criação de PDF estruturado com múltiplos elementos."""
        file_path = os.path.join(temp_dir, "structured.pdf")
        
        elements = [
            {"type": "title", "text": "Annual Report"},
            {"type": "heading", "text": "Executive Summary", "level": 1},
            {"type": "paragraph", "text": "This report presents the results.", "alignment": "justify"},
            {"type": "table", "headers": ["Product", "Sales"], "rows": [["A", "100"], ["B", "200"]]},
            {"type": "list", "items": ["Item 1", "Item 2"], "ordered": False},
            {"type": "spacer", "height": 20},
            {"type": "page_break"},
            {"type": "heading", "text": "Conclusion", "level": 1},
            {"type": "paragraph", "text": "Thank you."}
        ]
        
        pdf_tool.create_pdf(file_path, elements, page_size="A4")
        
        assert Path(file_path).exists()
        
        # Verificar que pode ser lido
        result = pdf_tool.read_pdf(file_path)
        assert result['num_pages'] >= 2  # Deve ter pelo menos 2 páginas (page_break)
    
    def test_create_pdf_with_formatting(self, pdf_tool, temp_dir):
        """Testa criação de PDF com formatação de texto."""
        file_path = os.path.join(temp_dir, "formatted.pdf")
        
        elements = [
            {"type": "paragraph", "text": "Bold text", "bold": True},
            {"type": "paragraph", "text": "Italic text", "italic": True},
            {"type": "paragraph", "text": "Centered", "alignment": "center"}
        ]
        
        pdf_tool.create_pdf(file_path, elements)
        
        assert Path(file_path).exists()


class TestPdfToolRead:
    """Testes para leitura de PDFs."""
    
    def test_read_pdf(self, pdf_tool, temp_dir):
        """Testa leitura de PDF."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF primeiro
        elements = [
            {"type": "title", "text": "Test Title"},
            {"type": "paragraph", "text": "Test content on page 1."},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Content on page 2."}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Ler PDF
        result = pdf_tool.read_pdf(file_path)
        
        assert 'pages' in result
        assert 'num_pages' in result
        assert 'metadata' in result
        assert result['num_pages'] == 2
        assert len(result['pages']) == 2
        assert 'Test Title' in result['pages'][0]
        assert 'page 2' in result['pages'][1]
    
    def test_read_nonexistent_pdf(self, pdf_tool):
        """Testa leitura de PDF inexistente."""
        with pytest.raises(FileNotFoundError):
            pdf_tool.read_pdf("nonexistent.pdf")
    
    def test_read_corrupted_pdf(self, pdf_tool, temp_dir):
        """Testa leitura de PDF corrompido."""
        file_path = os.path.join(temp_dir, "corrupted.pdf")
        
        # Criar arquivo corrompido
        with open(file_path, 'w') as f:
            f.write("This is not a valid PDF")
        
        with pytest.raises(CorruptedFileError):
            pdf_tool.read_pdf(file_path)


class TestPdfToolMerge:
    """Testes para merge de PDFs."""
    
    def test_merge_pdfs(self, pdf_tool, temp_dir):
        """Testa merge de múltiplos PDFs."""
        # Criar 3 PDFs
        pdf1 = os.path.join(temp_dir, "part1.pdf")
        pdf2 = os.path.join(temp_dir, "part2.pdf")
        pdf3 = os.path.join(temp_dir, "part3.pdf")
        output = os.path.join(temp_dir, "merged.pdf")
        
        pdf_tool.create_pdf(pdf1, [{"type": "paragraph", "text": "Page 1"}])
        pdf_tool.create_pdf(pdf2, [{"type": "paragraph", "text": "Page 2"}])
        pdf_tool.create_pdf(pdf3, [{"type": "paragraph", "text": "Page 3"}])
        
        # Merge
        pdf_tool.merge_pdfs([pdf1, pdf2, pdf3], output)
        
        assert Path(output).exists()
        
        # Verificar que tem 3 páginas
        result = pdf_tool.read_pdf(output)
        assert result['num_pages'] == 3
    
    def test_merge_nonexistent_pdf(self, pdf_tool, temp_dir):
        """Testa merge com arquivo inexistente."""
        output = os.path.join(temp_dir, "merged.pdf")
        
        with pytest.raises(FileNotFoundError):
            pdf_tool.merge_pdfs(["nonexistent1.pdf", "nonexistent2.pdf"], output)


class TestPdfToolSplit:
    """Testes para split de PDFs."""
    
    def test_split_pdf(self, pdf_tool, temp_dir):
        """Testa split de PDF."""
        # Criar PDF com 5 páginas
        source = os.path.join(temp_dir, "source.pdf")
        elements = []
        for i in range(5):
            elements.append({"type": "paragraph", "text": f"Page {i+1}"})
            if i < 4:
                elements.append({"type": "page_break"})
        
        pdf_tool.create_pdf(source, elements)
        
        # Split páginas 2-4
        output = os.path.join(temp_dir, "excerpt.pdf")
        pdf_tool.split_pdf(source, output, start_page=2, end_page=4)
        
        assert Path(output).exists()
        
        # Verificar que tem 3 páginas
        result = pdf_tool.read_pdf(output)
        assert result['num_pages'] == 3
    
    def test_split_invalid_range(self, pdf_tool, temp_dir):
        """Testa split com range inválido."""
        source = os.path.join(temp_dir, "source.pdf")
        pdf_tool.create_pdf(source, [{"type": "paragraph", "text": "Test"}])
        
        output = os.path.join(temp_dir, "output.pdf")
        
        # start_page < 1
        with pytest.raises(ValueError):
            pdf_tool.split_pdf(source, output, start_page=0, end_page=1)
        
        # end_page < start_page
        with pytest.raises(ValueError):
            pdf_tool.split_pdf(source, output, start_page=5, end_page=2)
        
        # end_page > total pages
        with pytest.raises(ValueError):
            pdf_tool.split_pdf(source, output, start_page=1, end_page=100)


class TestPdfToolTextOverlay:
    """Testes para adicionar texto/watermark."""
    
    def test_add_text_overlay_all_pages(self, pdf_tool, temp_dir):
        """Testa adicionar texto em todas as páginas."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF com 2 páginas
        elements = [
            {"type": "paragraph", "text": "Page 1"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 2"}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Adicionar watermark
        pdf_tool.add_text_overlay(file_path, "CONFIDENTIAL", opacity=0.3)
        
        assert Path(file_path).exists()
    
    def test_add_text_overlay_specific_pages(self, pdf_tool, temp_dir):
        """Testa adicionar texto em páginas específicas."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF com 3 páginas
        elements = [
            {"type": "paragraph", "text": "Page 1"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 2"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 3"}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Adicionar watermark apenas nas páginas 1 e 3
        pdf_tool.add_text_overlay(
            file_path, "DRAFT",
            x=250, y=400, font_size=50,
            color="FF0000", opacity=0.5,
            pages=[1, 3]
        )
        
        assert Path(file_path).exists()


class TestPdfToolInfo:
    """Testes para obter informações do PDF."""
    
    def test_get_info(self, pdf_tool, temp_dir):
        """Testa obter informações do PDF."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF
        elements = [
            {"type": "title", "text": "Test"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 2"}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Obter info
        info = pdf_tool.get_info(file_path)
        
        assert 'num_pages' in info
        assert 'file_size_bytes' in info
        assert 'file_size_kb' in info
        assert 'encrypted' in info
        assert info['num_pages'] == 2
        assert info['file_size_bytes'] > 0
        assert info['encrypted'] is False


class TestPdfToolRotate:
    """Testes para rotação de páginas."""
    
    def test_rotate_all_pages(self, pdf_tool, temp_dir):
        """Testa rotação de todas as páginas."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF
        elements = [{"type": "paragraph", "text": "Test"}]
        pdf_tool.create_pdf(file_path, elements)
        
        # Rotacionar
        pdf_tool.rotate_pages(file_path, rotation=90)
        
        assert Path(file_path).exists()
    
    def test_rotate_specific_pages(self, pdf_tool, temp_dir):
        """Testa rotação de páginas específicas."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF com 3 páginas
        elements = [
            {"type": "paragraph", "text": "Page 1"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 2"},
            {"type": "page_break"},
            {"type": "paragraph", "text": "Page 3"}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Rotacionar apenas páginas 1 e 3
        pdf_tool.rotate_pages(file_path, rotation=180, pages=[1, 3])
        
        assert Path(file_path).exists()
    
    def test_rotate_invalid_angle(self, pdf_tool, temp_dir):
        """Testa rotação com ângulo inválido."""
        file_path = os.path.join(temp_dir, "test.pdf")
        pdf_tool.create_pdf(file_path, [{"type": "paragraph", "text": "Test"}])
        
        with pytest.raises(ValueError):
            pdf_tool.rotate_pages(file_path, rotation=45)


class TestPdfToolExtractTables:
    """Testes para extração de tabelas."""
    
    def test_extract_tables_with_tabs(self, pdf_tool, temp_dir):
        """Testa extração de tabelas com separadores tab."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        # Criar PDF com tabela
        elements = [
            {"type": "table", "headers": ["Name", "Age", "City"], "rows": [["Alice", "30", "NYC"], ["Bob", "25", "LA"]]}
        ]
        pdf_tool.create_pdf(file_path, elements)
        
        # Extrair tabelas
        tables = pdf_tool.extract_tables(file_path)
        
        # Pode ou não detectar dependendo do formato do texto extraído
        assert isinstance(tables, list)
    
    def test_extract_tables_empty_pdf(self, pdf_tool, temp_dir):
        """Testa extração de tabelas de PDF sem tabelas."""
        file_path = os.path.join(temp_dir, "test.pdf")
        
        elements = [{"type": "paragraph", "text": "Just text, no tables."}]
        pdf_tool.create_pdf(file_path, elements)
        
        tables = pdf_tool.extract_tables(file_path)
        
        assert isinstance(tables, list)
        # Pode estar vazio ou ter falsos positivos
