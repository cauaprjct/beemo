"""Testes para o módulo security_validator."""

import os
import tempfile
from pathlib import Path

import pytest

from src.exceptions import ValidationError
from src.security_validator import (
    ALLOWED_OPERATIONS,
    SecurityValidator,
    sanitize_filename,
    validate_file_path,
    validate_operation,
)


class TestSecurityValidator:
    """Testes para a classe SecurityValidator."""
    
    @pytest.fixture
    def temp_root(self):
        """Cria diretório temporário para testes."""
        with tempfile.TemporaryDirectory() as tmpdir:
            yield tmpdir
    
    @pytest.fixture
    def validator(self, temp_root):
        """Cria instância do SecurityValidator."""
        return SecurityValidator(temp_root)
    
    def test_init(self, temp_root):
        """Testa inicialização do SecurityValidator."""
        validator = SecurityValidator(temp_root)
        assert validator.root_path == Path(temp_root).resolve()
    
    def test_validate_file_path_relative_within_root(self, validator, temp_root):
        """Testa validação de caminho relativo dentro do root_path."""
        file_path = "test.xlsx"
        result = validator.validate_file_path(file_path)
        
        expected = (Path(temp_root) / file_path).resolve()
        assert result == expected
    
    def test_validate_file_path_nested_within_root(self, validator, temp_root):
        """Testa validação de caminho aninhado dentro do root_path."""
        file_path = "subdir/nested/test.xlsx"
        result = validator.validate_file_path(file_path)
        
        expected = (Path(temp_root) / file_path).resolve()
        assert result == expected
    
    def test_validate_file_path_absolute_within_root(self, validator, temp_root):
        """Testa validação de caminho absoluto dentro do root_path."""
        file_path = os.path.join(temp_root, "test.xlsx")
        result = validator.validate_file_path(file_path)
        
        assert result == Path(file_path).resolve()
    
    def test_validate_file_path_traversal_attack(self, validator):
        """Testa detecção de path traversal com ../."""
        file_path = "../outside.xlsx"
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_path(file_path)
        
        assert "fora do diretório permitido" in str(exc_info.value)
    
    def test_validate_file_path_absolute_outside_root(self, validator):
        """Testa rejeição de caminho absoluto fora do root_path."""
        file_path = "/tmp/outside.xlsx"
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_path(file_path)
        
        assert "fora do diretório permitido" in str(exc_info.value)
    
    def test_validate_file_path_multiple_traversal(self, validator):
        """Testa detecção de múltiplas tentativas de path traversal."""
        file_path = "../../outside.xlsx"
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_path(file_path)
        
        assert "fora do diretório permitido" in str(exc_info.value)
    
    def test_validate_file_path_hidden_traversal(self, validator):
        """Testa detecção de path traversal escondido em subdiretórios."""
        file_path = "subdir/../../outside.xlsx"
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_path(file_path)
        
        assert "fora do diretório permitido" in str(exc_info.value)
    
    def test_is_within_root_path_true(self, validator, temp_root):
        """Testa verificação de arquivo dentro do root_path."""
        file_path = Path(temp_root) / "test.xlsx"
        assert validator._is_within_root_path(file_path) is True
    
    def test_is_within_root_path_false(self, validator):
        """Testa verificação de arquivo fora do root_path."""
        file_path = Path("/tmp/outside.xlsx")
        assert validator._is_within_root_path(file_path) is False
    
    def test_validate_operation_allowed(self, validator):
        """Testa validação de operações permitidas."""
        for operation in ALLOWED_OPERATIONS:
            # Não deve lançar exceção
            validator.validate_operation(operation)
    
    def test_validate_operation_not_allowed(self, validator):
        """Testa rejeição de operação não permitida."""
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_operation("delete")
        
        assert "Operação não permitida" in str(exc_info.value)
        assert "delete" in str(exc_info.value)
    
    def test_validate_operation_empty(self, validator):
        """Testa rejeição de operação vazia."""
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_operation("")
        
        assert "Operação não permitida" in str(exc_info.value)
    
    def test_sanitize_filename_clean(self, validator):
        """Testa sanitização de nome de arquivo limpo."""
        filename = "test_file.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == filename
    
    def test_sanitize_filename_with_spaces(self, validator):
        """Testa sanitização de nome com espaços."""
        filename = "test file.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "test file.xlsx"
    
    def test_sanitize_filename_with_special_chars(self, validator):
        """Testa remoção de caracteres especiais."""
        filename = "test@file#.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "testfile.xlsx"
    
    def test_sanitize_filename_with_path_separator(self, validator):
        """Testa remoção de separadores de caminho."""
        filename = "test/file.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "testfile.xlsx"
    
    def test_sanitize_filename_with_backslash(self, validator):
        """Testa remoção de backslash."""
        filename = "test\\file.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "testfile.xlsx"
    
    def test_sanitize_filename_with_double_dots(self, validator):
        """Testa remoção de pontos duplos."""
        filename = "test..file.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "test.file.xlsx"
    
    def test_sanitize_filename_with_leading_dots(self, validator):
        """Testa remoção de pontos no início."""
        filename = "...test.xlsx"
        result = validator.sanitize_filename(filename)
        assert result == "test.xlsx"
    
    def test_sanitize_filename_with_trailing_spaces(self, validator):
        """Testa remoção de espaços no início e fim."""
        filename = "  test.xlsx  "
        result = validator.sanitize_filename(filename)
        assert result == "test.xlsx"
    
    def test_sanitize_filename_empty_after_sanitization(self, validator):
        """Testa erro quando nome fica vazio após sanitização."""
        filename = "@#$%"
        
        with pytest.raises(ValidationError) as exc_info:
            validator.sanitize_filename(filename)
        
        assert "Nome de arquivo inválido" in str(exc_info.value)
    
    def test_sanitize_filename_only_dots(self, validator):
        """Testa erro quando nome contém apenas pontos."""
        filename = "..."
        
        with pytest.raises(ValidationError) as exc_info:
            validator.sanitize_filename(filename)
        
        assert "Nome de arquivo inválido" in str(exc_info.value)
    
    def test_validate_file_extension_allowed(self, validator):
        """Testa validação de extensão permitida."""
        allowed = {'.xlsx', '.docx', '.pptx'}
        
        # Não deve lançar exceção
        validator.validate_file_extension("test.xlsx", allowed)
        validator.validate_file_extension("test.docx", allowed)
        validator.validate_file_extension("test.pptx", allowed)
    
    def test_validate_file_extension_case_insensitive(self, validator):
        """Testa validação case-insensitive de extensão."""
        allowed = {'.xlsx'}
        
        # Não deve lançar exceção
        validator.validate_file_extension("test.XLSX", allowed)
        validator.validate_file_extension("test.XlSx", allowed)
    
    def test_validate_file_extension_not_allowed(self, validator):
        """Testa rejeição de extensão não permitida."""
        allowed = {'.xlsx'}
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_extension("test.txt", allowed)
        
        assert "Extensão de arquivo não permitida" in str(exc_info.value)
        assert ".txt" in str(exc_info.value)
    
    def test_validate_file_extension_no_extension(self, validator):
        """Testa rejeição de arquivo sem extensão."""
        allowed = {'.xlsx'}
        
        with pytest.raises(ValidationError) as exc_info:
            validator.validate_file_extension("test", allowed)
        
        assert "Extensão de arquivo não permitida" in str(exc_info.value)


class TestHelperFunctions:
    """Testes para funções auxiliares."""
    
    def test_validate_file_path_helper(self):
        """Testa função auxiliar validate_file_path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = "test.xlsx"
            result = validate_file_path(file_path, tmpdir)
            
            expected = (Path(tmpdir) / file_path).resolve()
            assert result == expected
    
    def test_validate_file_path_helper_outside_root(self):
        """Testa função auxiliar com caminho fora do root."""
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = "../outside.xlsx"
            
            with pytest.raises(ValidationError):
                validate_file_path(file_path, tmpdir)
    
    def test_validate_operation_helper_allowed(self):
        """Testa função auxiliar validate_operation com operação permitida."""
        # Não deve lançar exceção
        validate_operation("read")
        validate_operation("create")
        validate_operation("update")
    
    def test_validate_operation_helper_not_allowed(self):
        """Testa função auxiliar validate_operation com operação não permitida."""
        with pytest.raises(ValidationError) as exc_info:
            validate_operation("delete")
        
        assert "Operação não permitida" in str(exc_info.value)
    
    def test_sanitize_filename_helper(self):
        """Testa função auxiliar sanitize_filename."""
        result = sanitize_filename("test@file#.xlsx")
        assert result == "testfile.xlsx"
    
    def test_sanitize_filename_helper_invalid(self):
        """Testa função auxiliar com nome inválido."""
        with pytest.raises(ValidationError):
            sanitize_filename("@#$%")
