"""Módulo de gerenciamento de versões de arquivos.

Este módulo fornece funcionalidade de undo/redo através de versionamento
de arquivos, permitindo que usuários desfaçam e refaçam operações.
"""

import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict

from src.exceptions import ValidationError
from src.logging_config import get_logger

logger = get_logger(__name__)


@dataclass
class Version:
    """Representa uma versão de um arquivo."""
    version_id: str
    timestamp: str
    operation: str
    backup_path: str  # Estado ANTES da operação
    user_prompt: str
    file_size: int
    after_path: Optional[str] = None  # Estado APÓS a operação (para redo)


@dataclass
class FileHistory:
    """Histórico de versões de um arquivo."""
    file_path: str
    versions: List[Version]
    current_index: int  # -1 significa estado atual (sem undo)


class VersionManager:
    """Gerenciador de versões de arquivos com suporte a undo/redo."""
    
    def __init__(self, root_path: str, max_versions: int = 10):
        """
        Inicializa o gerenciador de versões.
        
        Args:
            root_path: Caminho raiz onde os arquivos são armazenados
            max_versions: Número máximo de versões por arquivo
        """
        self.root_path = Path(root_path).resolve()
        self.max_versions = max_versions
        self.versions_dir = self.root_path / ".versions"
        self.metadata_file = self.versions_dir / "metadata.json"
        
        # Criar diretório de versões se não existir (com parents=True)
        self.versions_dir.mkdir(parents=True, exist_ok=True)
        
        # Carregar ou inicializar metadata
        self.metadata: Dict[str, FileHistory] = self._load_metadata()
        
        logger.info(
            f"VersionManager inicializado: root={self.root_path}, "
            f"max_versions={self.max_versions}"
        )
    
    def create_backup(
        self,
        file_path: str,
        operation: str,
        user_prompt: str
    ) -> Optional[str]:
        """
        Cria backup de um arquivo antes de modificação.
        
        Args:
            file_path: Caminho do arquivo a fazer backup
            operation: Tipo de operação (update, add_sheet, etc.)
            user_prompt: Prompt do usuário que gerou a operação
        
        Returns:
            version_id do backup criado, ou None se arquivo não existe
        """
        file_path_obj = Path(file_path).resolve()
        
        # Se arquivo não existe, não há o que fazer backup
        if not file_path_obj.exists():
            logger.debug(f"Arquivo não existe, sem backup: {file_path}")
            return None
        
        # Gerar ID único para esta versão
        version_id = str(uuid.uuid4())
        
        # Criar caminho do backup (estado ANTES)
        backup_filename = f"{file_path_obj.stem}_{version_id}_before{file_path_obj.suffix}"
        backup_path = self.versions_dir / backup_filename
        
        # Copiar arquivo (estado antes da operação)
        shutil.copy2(file_path_obj, backup_path)
        
        # Criar objeto Version
        version = Version(
            version_id=version_id,
            timestamp=datetime.now().isoformat(),
            operation=operation,
            backup_path=str(backup_path),
            user_prompt=user_prompt,
            file_size=file_path_obj.stat().st_size,
            after_path=None  # Será preenchido depois da operação
        )
        
        # Obter ou criar histórico do arquivo
        rel_path = str(file_path_obj.relative_to(self.root_path))
        if rel_path not in self.metadata:
            self.metadata[rel_path] = FileHistory(
                file_path=rel_path,
                versions=[],
                current_index=-1
            )
        
        history = self.metadata[rel_path]
        
        # Se estamos em um estado de undo, descartar versões "futuras"
        if history.current_index < len(history.versions) - 1:
            discarded = history.versions[history.current_index + 1:]
            history.versions = history.versions[:history.current_index + 1]
            
            # Remover arquivos de backup descartados
            for v in discarded:
                self._remove_version_files(v)
        
        # Adicionar nova versão
        history.versions.append(version)
        history.current_index = len(history.versions) - 1
        
        # Limpar versões antigas se exceder o limite
        self._cleanup_old_versions(rel_path)
        
        # Salvar metadata
        self._save_metadata()
        
        logger.info(
            f"Backup criado: {rel_path} -> {backup_filename} "
            f"(operação: {operation})"
        )
        
        return version_id
    
    def save_after_state(self, file_path: str, version_id: str) -> None:
        """
        Salva o estado do arquivo APÓS uma operação (para suportar redo).
        
        Args:
            file_path: Caminho do arquivo
            version_id: ID da versão criada por create_backup
        """
        file_path_obj = Path(file_path).resolve()
        rel_path = str(file_path_obj.relative_to(self.root_path))
        
        if rel_path not in self.metadata:
            logger.warning(f"Arquivo sem histórico: {rel_path}")
            return
        
        history = self.metadata[rel_path]
        
        # Encontrar a versão pelo ID
        version = None
        for v in history.versions:
            if v.version_id == version_id:
                version = v
                break
        
        if not version:
            logger.warning(f"Versão não encontrada: {version_id}")
            return
        
        # Criar caminho do backup (estado APÓS)
        after_filename = f"{file_path_obj.stem}_{version_id}_after{file_path_obj.suffix}"
        after_path = self.versions_dir / after_filename
        
        # Copiar arquivo (estado após a operação)
        if file_path_obj.exists():
            shutil.copy2(file_path_obj, after_path)
            version.after_path = str(after_path)
            
            # Salvar metadata
            self._save_metadata()
            
            logger.debug(f"Estado pós-operação salvo: {after_filename}")


    
    def undo(self, file_path: str) -> bool:
        """
        Desfaz a última operação em um arquivo.
        
        Args:
            file_path: Caminho do arquivo
        
        Returns:
            True se undo foi realizado, False caso contrário
        
        Raises:
            ValidationError: Se não há versões para desfazer
        """
        file_path_obj = Path(file_path).resolve()
        rel_path = str(file_path_obj.relative_to(self.root_path))
        
        if not self.can_undo(file_path):
            raise ValidationError(f"Não há versões para desfazer: {rel_path}")
        
        history = self.metadata[rel_path]
        
        # Restaurar versão anterior
        version = history.versions[history.current_index]
        backup_path = Path(version.backup_path)
        
        if not backup_path.exists():
            raise ValidationError(
                f"Arquivo de backup não encontrado: {backup_path}"
            )
        
        # Copiar backup de volta para o arquivo original
        shutil.copy2(backup_path, file_path_obj)
        
        # Decrementar índice
        history.current_index -= 1
        
        # Salvar metadata
        self._save_metadata()
        
        logger.info(
            f"Undo realizado: {rel_path} -> versão {version.version_id}"
        )
        
        return True
    
    def redo(self, file_path: str) -> bool:
        """
        Refaz uma operação desfeita em um arquivo.
        
        Args:
            file_path: Caminho do arquivo
        
        Returns:
            True se redo foi realizado, False caso contrário
        
        Raises:
            ValidationError: Se não há versões para refazer
        """
        file_path_obj = Path(file_path).resolve()
        rel_path = str(file_path_obj.relative_to(self.root_path))
        
        if not self.can_redo(file_path):
            raise ValidationError(f"Não há versões para refazer: {rel_path}")
        
        history = self.metadata[rel_path]
        
        # Incrementar índice para próxima versão
        history.current_index += 1
        
        # Restaurar estado APÓS a operação
        version = history.versions[history.current_index]
        
        if version.after_path and Path(version.after_path).exists():
            shutil.copy2(version.after_path, file_path_obj)
            logger.info(
                f"Redo realizado: {rel_path} -> versão {version.version_id}"
            )
        else:
            logger.warning(
                f"Estado pós-operação não disponível para {version.version_id}"
            )
        
        # Salvar metadata
        self._save_metadata()
        
        return True
    
    def can_undo(self, file_path: str) -> bool:
        """
        Verifica se há versões disponíveis para desfazer.
        
        Args:
            file_path: Caminho do arquivo
        
        Returns:
            True se undo está disponível
        """
        try:
            file_path_obj = Path(file_path).resolve()
            rel_path = str(file_path_obj.relative_to(self.root_path))
            
            if rel_path not in self.metadata:
                return False
            
            history = self.metadata[rel_path]
            return history.current_index >= 0
        except Exception:
            return False
    
    def can_redo(self, file_path: str) -> bool:
        """
        Verifica se há versões disponíveis para refazer.
        
        Args:
            file_path: Caminho do arquivo
        
        Returns:
            True se redo está disponível
        """
        try:
            file_path_obj = Path(file_path).resolve()
            rel_path = str(file_path_obj.relative_to(self.root_path))
            
            if rel_path not in self.metadata:
                return False
            
            history = self.metadata[rel_path]
            return history.current_index < len(history.versions) - 1
        except Exception:
            return False
    
    def get_history(self, file_path: str) -> Optional[FileHistory]:
        """
        Retorna o histórico de versões de um arquivo.
        
        Args:
            file_path: Caminho do arquivo
        
        Returns:
            FileHistory ou None se não há histórico
        """
        try:
            file_path_obj = Path(file_path).resolve()
            rel_path = str(file_path_obj.relative_to(self.root_path))
            return self.metadata.get(rel_path)
        except Exception:
            return None
    
    def get_all_files_with_history(self) -> List[str]:
        """
        Retorna lista de todos os arquivos que têm histórico de versões.
        
        Returns:
            Lista de caminhos relativos de arquivos
        """
        return list(self.metadata.keys())
    
    def _cleanup_old_versions(self, rel_path: str) -> None:
        """
        Remove versões antigas se exceder o limite.
        
        Args:
            rel_path: Caminho relativo do arquivo
        """
        history = self.metadata[rel_path]
        
        if len(history.versions) > self.max_versions:
            # Remover versões mais antigas
            to_remove = history.versions[:len(history.versions) - self.max_versions]
            history.versions = history.versions[-self.max_versions:]
            
            # Ajustar current_index
            history.current_index = min(
                history.current_index,
                len(history.versions) - 1
            )
            
            # Remover arquivos de backup
            for version in to_remove:
                self._remove_version_files(version)
    
    def _remove_version_files(self, version: Version) -> None:
        """
        Remove arquivos de backup de uma versão.
        
        Args:
            version: Versão a remover
        """
        try:
            Path(version.backup_path).unlink(missing_ok=True)
            logger.debug(f"Backup removido: {version.backup_path}")
        except Exception as e:
            logger.warning(f"Erro ao remover backup: {e}")
        
        if version.after_path:
            try:
                Path(version.after_path).unlink(missing_ok=True)
                logger.debug(f"Estado pós-operação removido: {version.after_path}")
            except Exception as e:
                logger.warning(f"Erro ao remover estado pós-operação: {e}")
    
    def _load_metadata(self) -> Dict[str, FileHistory]:
        """
        Carrega metadata do arquivo JSON.
        
        Returns:
            Dicionário de históricos de arquivos
        """
        if not self.metadata_file.exists():
            return {}
        
        try:
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Converter dicionários em objetos
            metadata = {}
            for rel_path, history_dict in data.items():
                versions = [
                    Version(**v) for v in history_dict['versions']
                ]
                metadata[rel_path] = FileHistory(
                    file_path=history_dict['file_path'],
                    versions=versions,
                    current_index=history_dict['current_index']
                )
            
            logger.debug(f"Metadata carregado: {len(metadata)} arquivos")
            return metadata
        except Exception as e:
            logger.error(f"Erro ao carregar metadata: {e}", exc_info=True)
            return {}
    
    def _save_metadata(self) -> None:
        """Salva metadata no arquivo JSON."""
        try:
            # Converter objetos em dicionários
            data = {}
            for rel_path, history in self.metadata.items():
                data[rel_path] = {
                    'file_path': history.file_path,
                    'versions': [asdict(v) for v in history.versions],
                    'current_index': history.current_index
                }
            
            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.debug("Metadata salvo com sucesso")
        except Exception as e:
            logger.error(f"Erro ao salvar metadata: {e}", exc_info=True)
