"""File scanner para identificar arquivos Office em diretórios locais."""

import logging
from pathlib import Path
from typing import List

from src.logging_config import get_logger

logger = get_logger(__name__)


class FileScanner:
    """Scanner para varrer diretórios e identificar arquivos Office."""
    
    # Extensões válidas de arquivos Office
    OFFICE_EXTENSIONS = {'.xlsx', '.docx', '.pptx', '.pdf'}
    
    # Prefixo de arquivos temporários do Office
    TEMP_FILE_PREFIX = '~$'
    
    # Diretórios a ignorar durante a varredura
    IGNORED_DIRS = {'.versions', '.cache'}
    
    def __init__(self, root_path: str):
        """
        Inicializa scanner com pasta raiz.
        
        Args:
            root_path: Caminho da pasta raiz para varredura
        """
        self.root_path = Path(root_path)
        logger.info(f"FileScanner inicializado com pasta raiz: {self.root_path}")
    
    def scan_office_files(self) -> List[str]:
        """
        Retorna lista de caminhos de arquivos Office encontrados.
        
        Varre recursivamente a pasta raiz e identifica todos os arquivos
        com extensões .xlsx, .docx e .pptx, excluindo arquivos temporários.
        
        Returns:
            Lista de caminhos completos dos arquivos Office encontrados
        """
        # Verifica se a pasta raiz existe
        if not self.root_path.exists():
            logger.warning(f"Pasta raiz não existe: {self.root_path}")
            return []
        
        if not self.root_path.is_dir():
            logger.warning(f"Caminho não é um diretório: {self.root_path}")
            return []
        
        office_files = []
        
        try:
            # Varre recursivamente todos os arquivos
            for file_path in self.root_path.rglob('*'):
                # Ignora diretórios
                if not file_path.is_file():
                    continue
                
                # Ignora arquivos em diretórios internos (.versions, .cache)
                if any(part in self.IGNORED_DIRS for part in file_path.parts):
                    continue
                
                # Verifica se é arquivo Office válido
                if self._is_office_file(file_path) and not self._is_temp_file(file_path.name):
                    office_files.append(str(file_path.absolute()))
                    logger.debug(f"Arquivo Office encontrado: {file_path}")
            
            logger.info(f"Varredura concluída: {len(office_files)} arquivo(s) encontrado(s)")
            
        except PermissionError as e:
            logger.error(f"Erro de permissão ao acessar diretório: {e}")
        except Exception as e:
            logger.error(f"Erro durante varredura de arquivos: {e}", exc_info=True)
        
        return office_files
    
    def _is_office_file(self, file_path: Path) -> bool:
        """
        Verifica se arquivo é Office válido.
        
        Args:
            file_path: Caminho do arquivo a verificar
        
        Returns:
            True se o arquivo tem extensão Office válida
        """
        return file_path.suffix.lower() in self.OFFICE_EXTENSIONS
    
    def _is_temp_file(self, file_name: str) -> bool:
        """
        Verifica se arquivo é temporário (~$).
        
        Arquivos temporários do Office começam com ~$ e devem ser ignorados.
        
        Args:
            file_name: Nome do arquivo a verificar
        
        Returns:
            True se o arquivo é temporário
        """
        return file_name.startswith(self.TEMP_FILE_PREFIX)
