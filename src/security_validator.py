"""Módulo de validação de segurança para operações de arquivo.

Este módulo fornece funções para validar operações de arquivo e prevenir
ataques de segurança como path traversal, garantindo que todas as operações
ocorram dentro do diretório raiz configurado.
"""

import logging
import os
import re
from pathlib import Path
from typing import List, Set

from src.exceptions import ValidationError
from src.logging_config import get_logger

logger = get_logger(__name__)


# Whitelist de operações permitidas
ALLOWED_OPERATIONS = {
    'read',
    'create',
    'update',
    'add',
    'append',
    'delete_sheet',
    'delete_rows',
    'format',
    'auto_width',
    'formula',
    'merge',
    'add_chart',
    'update_chart',
    'move_chart',
    'resize_chart',
    'list_charts',
    'delete_chart',
    'sort',
    'remove_duplicates',
    'filter_and_copy',
    'insert_rows',
    'insert_columns',
    'freeze_panes',
    'unfreeze_panes',
    'find_and_replace',
    'add_heading',
    'add_table',
    'add_list',
    'delete_paragraph',
    'replace',
    'add_sheet',
    'add_paragraph',
    'add_slide',
    'delete_slide',
    'duplicate_slide',
    'add_textbox',
    'update_cell',
    'update_paragraph',
    'update_slide',
    'extract_tables',
    'extract_text',
    'split',
    'add_text',
    'get_info',
    'rotate',
    'improve_text',
    'correct_grammar',
    'improve_clarity',
    'adjust_tone',
    'simplify_language',
    'rewrite_professional',
    'format_paragraph',
    'generate_summary',
    'extract_key_points',
    'create_resume',
    'generate_conclusions',
    'create_faq',
    'convert_list_to_table',
    'convert_table_to_list',
    'extract_tables_to_excel',
    'export_to_txt',
    'export_to_markdown',
    'export_to_html',
    'export_to_pdf',
    'analyze_word_count',
    'analyze_section_length',
    'get_document_statistics',
    'analyze_tone',
    'identify_jargon',
    'analyze_readability',
    'check_term_consistency',
    'analyze_document',
    'add_image',
    'add_image_at_position',
    'add_hyperlink',
    'add_hyperlink_to_paragraph',
    'add_header',
    'add_footer',
    'remove_header',
    'remove_footer',
    'set_page_margins',
    'set_page_size',
    'get_page_layout',
    'add_page_break',
    'add_section_break'
}


class SecurityValidator:
    """Validador de segurança para operações de arquivo."""
    
    def __init__(self, root_path: str):
        """
        Inicializa validador com pasta raiz.
        
        Args:
            root_path: Caminho da pasta raiz permitida para operações
        """
        self.root_path = Path(root_path).resolve()
        logger.info(f"SecurityValidator inicializado com root_path: {self.root_path}")
    
    def validate_file_path(self, file_path: str) -> Path:
        """
        Valida caminho de arquivo e previne path traversal.
        
        Verifica que o arquivo está dentro do root_path e não contém
        sequências perigosas como ../ ou caminhos absolutos suspeitos.
        
        Args:
            file_path: Caminho do arquivo a validar
        
        Returns:
            Path: Caminho resolvido e validado
        
        Raises:
            ValidationError: Se o caminho for inválido ou inseguro
        """
        try:
            # Converte para Path e resolve para caminho absoluto
            path = Path(file_path)
            
            # Se for caminho relativo, resolve em relação ao root_path
            if not path.is_absolute():
                resolved_path = (self.root_path / path).resolve()
            else:
                resolved_path = path.resolve()
            
            # Verifica se o caminho está dentro do root_path
            if not self._is_within_root_path(resolved_path):
                logger.warning(
                    f"Tentativa de acesso fora do root_path: {file_path} -> {resolved_path}"
                )
                raise ValidationError(
                    f"Caminho de arquivo fora do diretório permitido: {file_path}"
                )
            
            logger.debug(f"Caminho validado: {file_path} -> {resolved_path}")
            return resolved_path
            
        except (ValueError, OSError) as e:
            logger.error(f"Erro ao validar caminho: {file_path} - {e}")
            raise ValidationError(f"Caminho de arquivo inválido: {file_path}") from e
    
    def _is_within_root_path(self, file_path: Path) -> bool:
        """
        Verifica se arquivo está dentro do root_path.
        
        Args:
            file_path: Caminho resolvido do arquivo
        
        Returns:
            True se o arquivo está dentro do root_path
        """
        try:
            # Verifica se o caminho começa com root_path
            file_path.relative_to(self.root_path)
            return True
        except ValueError:
            return False
    
    def validate_operation(self, operation: str) -> None:
        """
        Valida se operação está na whitelist de operações permitidas.
        
        Args:
            operation: Nome da operação a validar
        
        Raises:
            ValidationError: Se a operação não for permitida
        """
        if operation not in ALLOWED_OPERATIONS:
            logger.warning(f"Tentativa de operação não permitida: {operation}")
            raise ValidationError(
                f"Operação não permitida: {operation}. "
                f"Operações permitidas: {', '.join(sorted(ALLOWED_OPERATIONS))}"
            )
        
        logger.debug(f"Operação validada: {operation}")
    
    def sanitize_filename(self, filename: str) -> str:
        """
        Sanitiza nome de arquivo removendo caracteres perigosos.
        
        Remove ou substitui caracteres que podem ser usados para
        path traversal ou outros ataques.
        
        Args:
            filename: Nome do arquivo a sanitizar
        
        Returns:
            Nome de arquivo sanitizado
        
        Raises:
            ValidationError: Se o nome do arquivo for vazio após sanitização
        """
        # Remove path separators e caracteres especiais perigosos
        # Mantém apenas letras, números, pontos, hífens, underscores e espaços
        sanitized = re.sub(r'[^\w\s\-.]', '', filename)
        
        # Remove múltiplos pontos consecutivos (previne ../)
        sanitized = re.sub(r'\.{2,}', '.', sanitized)
        
        # Remove espaços no início e fim
        sanitized = sanitized.strip()
        
        # Remove pontos no início (previne arquivos ocultos não intencionais)
        sanitized = sanitized.lstrip('.')
        
        if not sanitized:
            logger.error(f"Nome de arquivo vazio após sanitização: {filename}")
            raise ValidationError(
                f"Nome de arquivo inválido: {filename}"
            )
        
        if sanitized != filename:
            logger.info(f"Nome de arquivo sanitizado: '{filename}' -> '{sanitized}'")
        
        return sanitized
    
    def validate_file_extension(self, file_path: str, allowed_extensions: Set[str]) -> None:
        """
        Valida extensão do arquivo.
        
        Args:
            file_path: Caminho do arquivo
            allowed_extensions: Conjunto de extensões permitidas (ex: {'.xlsx', '.docx'})
        
        Raises:
            ValidationError: Se a extensão não for permitida
        """
        path = Path(file_path)
        extension = path.suffix.lower()
        
        if extension not in allowed_extensions:
            logger.warning(
                f"Extensão não permitida: {extension} em {file_path}"
            )
            raise ValidationError(
                f"Extensão de arquivo não permitida: {extension}. "
                f"Extensões permitidas: {', '.join(sorted(allowed_extensions))}"
            )
        
        logger.debug(f"Extensão validada: {extension} para {file_path}")


def validate_file_path(file_path: str, root_path: str) -> Path:
    """
    Função auxiliar para validar caminho de arquivo.
    
    Args:
        file_path: Caminho do arquivo a validar
        root_path: Caminho da pasta raiz permitida
    
    Returns:
        Path: Caminho resolvido e validado
    
    Raises:
        ValidationError: Se o caminho for inválido ou inseguro
    """
    validator = SecurityValidator(root_path)
    return validator.validate_file_path(file_path)


def validate_operation(operation: str) -> None:
    """
    Função auxiliar para validar operação.
    
    Args:
        operation: Nome da operação a validar
    
    Raises:
        ValidationError: Se a operação não for permitida
    """
    if operation not in ALLOWED_OPERATIONS:
        raise ValidationError(
            f"Operação não permitida: {operation}. "
            f"Operações permitidas: {', '.join(sorted(ALLOWED_OPERATIONS))}"
        )


def sanitize_filename(filename: str) -> str:
    """
    Função auxiliar para sanitizar nome de arquivo.
    
    Args:
        filename: Nome do arquivo a sanitizar
    
    Returns:
        Nome de arquivo sanitizado
    
    Raises:
        ValidationError: Se o nome do arquivo for vazio após sanitização
    """
    validator = SecurityValidator(os.getcwd())
    return validator.sanitize_filename(filename)
