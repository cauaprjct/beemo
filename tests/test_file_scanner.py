"""Testes para o módulo FileScanner."""

import pytest
from pathlib import Path
import tempfile
import os

from src.file_scanner import FileScanner


class TestFileScanner:
    """Testes unitários para a classe FileScanner."""
    
    def test_init(self):
        """Testa inicialização do FileScanner."""
        scanner = FileScanner("/test/path")
        assert scanner.root_path == Path("/test/path")
    
    def test_is_office_file_xlsx(self):
        """Testa identificação de arquivo Excel."""
        scanner = FileScanner(".")
        assert scanner._is_office_file(Path("test.xlsx")) is True
        assert scanner._is_office_file(Path("test.XLSX")) is True
    
    def test_is_office_file_docx(self):
        """Testa identificação de arquivo Word."""
        scanner = FileScanner(".")
        assert scanner._is_office_file(Path("test.docx")) is True
        assert scanner._is_office_file(Path("test.DOCX")) is True
    
    def test_is_office_file_pptx(self):
        """Testa identificação de arquivo PowerPoint."""
        scanner = FileScanner(".")
        assert scanner._is_office_file(Path("test.pptx")) is True
        assert scanner._is_office_file(Path("test.PPTX")) is True
    
    def test_is_office_file_invalid(self):
        """Testa rejeição de arquivos não-Office."""
        scanner = FileScanner(".")
        assert scanner._is_office_file(Path("test.txt")) is False
        assert scanner._is_office_file(Path("test.xls")) is False
        assert scanner._is_office_file(Path("test.doc")) is False
    
    def test_is_office_file_pdf(self):
        """Testa aceitação de arquivos PDF."""
        scanner = FileScanner(".")
        assert scanner._is_office_file(Path("test.pdf")) is True
        assert scanner._is_office_file(Path("test.PDF")) is True
    
    def test_is_temp_file(self):
        """Testa identificação de arquivos temporários."""
        scanner = FileScanner(".")
        assert scanner._is_temp_file("~$test.xlsx") is True
        assert scanner._is_temp_file("~$document.docx") is True
        assert scanner._is_temp_file("test.xlsx") is False
        assert scanner._is_temp_file("document.docx") is False
    
    def test_scan_nonexistent_directory(self):
        """Testa varredura de diretório inexistente."""
        scanner = FileScanner("/nonexistent/path/12345")
        result = scanner.scan_office_files()
        assert result == []
    
    def test_scan_empty_directory(self):
        """Testa varredura de diretório vazio."""
        with tempfile.TemporaryDirectory() as tmpdir:
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            assert result == []
    
    def test_scan_with_office_files(self):
        """Testa varredura de diretório com arquivos Office."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Cria arquivos de teste
            xlsx_file = Path(tmpdir) / "test.xlsx"
            docx_file = Path(tmpdir) / "test.docx"
            pptx_file = Path(tmpdir) / "test.pptx"
            
            xlsx_file.touch()
            docx_file.touch()
            pptx_file.touch()
            
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            
            assert len(result) == 3
            assert any("test.xlsx" in path for path in result)
            assert any("test.docx" in path for path in result)
            assert any("test.pptx" in path for path in result)
    
    def test_scan_excludes_temp_files(self):
        """Testa exclusão de arquivos temporários."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Cria arquivos normais e temporários
            normal_file = Path(tmpdir) / "test.xlsx"
            temp_file = Path(tmpdir) / "~$test.xlsx"
            
            normal_file.touch()
            temp_file.touch()
            
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            
            assert len(result) == 1
            assert "test.xlsx" in result[0]
            assert "~$test.xlsx" not in result[0]
    
    def test_scan_excludes_non_office_files(self):
        """Testa exclusão de arquivos não-Office."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Cria arquivos Office e não-Office
            xlsx_file = Path(tmpdir) / "test.xlsx"
            txt_file = Path(tmpdir) / "test.txt"
            pdf_file = Path(tmpdir) / "test.pdf"
            
            xlsx_file.touch()
            txt_file.touch()
            pdf_file.touch()
            
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            
            # Agora PDF é considerado arquivo Office
            assert len(result) == 2
            assert any("test.xlsx" in r for r in result)
            assert any("test.pdf" in r for r in result)
            assert not any("test.txt" in r for r in result)
    
    def test_scan_recursive(self):
        """Testa varredura recursiva de subdiretórios."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Cria estrutura de diretórios
            subdir1 = Path(tmpdir) / "subdir1"
            subdir2 = Path(tmpdir) / "subdir1" / "subdir2"
            subdir1.mkdir()
            subdir2.mkdir()
            
            # Cria arquivos em diferentes níveis
            file1 = Path(tmpdir) / "root.xlsx"
            file2 = subdir1 / "level1.docx"
            file3 = subdir2 / "level2.pptx"
            
            file1.touch()
            file2.touch()
            file3.touch()
            
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            
            assert len(result) == 3
            assert any("root.xlsx" in path for path in result)
            assert any("level1.docx" in path for path in result)
            assert any("level2.pptx" in path for path in result)
    
    def test_scan_returns_absolute_paths(self):
        """Testa que caminhos retornados são absolutos."""
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_file = Path(tmpdir) / "test.xlsx"
            xlsx_file.touch()
            
            scanner = FileScanner(tmpdir)
            result = scanner.scan_office_files()
            
            assert len(result) == 1
            assert Path(result[0]).is_absolute()
    
    def test_scan_file_instead_of_directory(self):
        """Testa comportamento quando root_path é um arquivo."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = Path(tmpdir) / "test.txt"
            file_path.touch()
            
            scanner = FileScanner(str(file_path))
            result = scanner.scan_office_files()
            
            assert result == []
