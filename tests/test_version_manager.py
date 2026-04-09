"""Testes para o módulo version_manager."""

import pytest
import tempfile
import shutil
from pathlib import Path

from src.version_manager import VersionManager, Version, FileHistory
from src.exceptions import ValidationError


@pytest.fixture
def temp_dir():
    """Cria diretório temporário para testes."""
    temp_path = tempfile.mkdtemp()
    yield temp_path
    shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def version_manager(temp_dir):
    """Cria VersionManager para testes."""
    return VersionManager(temp_dir, max_versions=5)


@pytest.fixture
def sample_file(temp_dir):
    """Cria arquivo de exemplo para testes."""
    file_path = Path(temp_dir) / "test.txt"
    file_path.write_text("Conteúdo inicial")
    return str(file_path)


def test_version_manager_initialization(temp_dir):
    """Testa inicialização do VersionManager."""
    vm = VersionManager(temp_dir, max_versions=10)
    
    assert vm.root_path == Path(temp_dir).resolve()
    assert vm.max_versions == 10
    assert vm.versions_dir.exists()
    # metadata.json é criado apenas quando há dados
    assert vm.metadata == {}


def test_create_backup(version_manager, sample_file):
    """Testa criação de backup."""
    version_id = version_manager.create_backup(
        sample_file,
        "update",
        "Teste de backup"
    )
    
    assert version_id is not None
    
    # Verificar que backup foi criado
    history = version_manager.get_history(sample_file)
    assert history is not None
    assert len(history.versions) == 1
    assert history.versions[0].operation == "update"
    assert history.current_index == 0


def test_create_backup_nonexistent_file(version_manager, temp_dir):
    """Testa backup de arquivo inexistente."""
    nonexistent = str(Path(temp_dir) / "nao_existe.txt")
    
    version_id = version_manager.create_backup(
        nonexistent,
        "update",
        "Teste"
    )
    
    assert version_id is None


def test_undo(version_manager, sample_file):
    """Testa operação de undo."""
    # Criar backup do estado inicial
    version_manager.create_backup(sample_file, "update", "Backup inicial")
    
    # Modificar arquivo
    Path(sample_file).write_text("Conteúdo modificado")
    
    # Fazer undo
    result = version_manager.undo(sample_file)
    assert result is True
    
    # Verificar que arquivo foi restaurado
    content = Path(sample_file).read_text()
    assert content == "Conteúdo inicial"


def test_undo_without_history(version_manager, sample_file):
    """Testa undo sem histórico."""
    with pytest.raises(ValidationError):
        version_manager.undo(sample_file)


def test_can_undo(version_manager, sample_file):
    """Testa verificação de disponibilidade de undo."""
    assert not version_manager.can_undo(sample_file)
    
    version_manager.create_backup(sample_file, "update", "Teste")
    assert version_manager.can_undo(sample_file)
    
    version_manager.undo(sample_file)
    assert not version_manager.can_undo(sample_file)


def test_can_redo(version_manager, sample_file):
    """Testa verificação de disponibilidade de redo."""
    assert not version_manager.can_redo(sample_file)
    
    version_manager.create_backup(sample_file, "update", "Teste")
    assert not version_manager.can_redo(sample_file)
    
    version_manager.undo(sample_file)
    assert version_manager.can_redo(sample_file)


def test_multiple_versions(version_manager, sample_file):
    """Testa múltiplas versões."""
    # Criar 3 versões
    for i in range(3):
        version_manager.create_backup(sample_file, "update", f"Versão {i}")
        Path(sample_file).write_text(f"Conteúdo {i}")
    
    history = version_manager.get_history(sample_file)
    assert len(history.versions) == 3
    assert history.current_index == 2


def test_cleanup_old_versions(version_manager, sample_file):
    """Testa limpeza de versões antigas."""
    # Criar mais versões que o limite (max_versions=5)
    for i in range(7):
        version_manager.create_backup(sample_file, "update", f"Versão {i}")
        Path(sample_file).write_text(f"Conteúdo {i}")
    
    history = version_manager.get_history(sample_file)
    assert len(history.versions) == 5  # Deve manter apenas 5


def test_discard_future_versions(version_manager, sample_file):
    """Testa descarte de versões futuras após nova operação."""
    # Criar 3 versões
    for i in range(3):
        version_manager.create_backup(sample_file, "update", f"Versão {i}")
        Path(sample_file).write_text(f"Conteúdo {i}")
    
    # Fazer undo 2 vezes
    version_manager.undo(sample_file)
    version_manager.undo(sample_file)
    
    history = version_manager.get_history(sample_file)
    assert history.current_index == 0
    
    # Criar nova versão (deve descartar versões futuras)
    version_manager.create_backup(sample_file, "update", "Nova versão")
    
    history = version_manager.get_history(sample_file)
    assert len(history.versions) == 2  # Versão 0 + nova versão
    assert not version_manager.can_redo(sample_file)


def test_get_all_files_with_history(version_manager, temp_dir):
    """Testa listagem de arquivos com histórico."""
    # Criar múltiplos arquivos
    file1 = Path(temp_dir) / "file1.txt"
    file2 = Path(temp_dir) / "file2.txt"
    
    file1.write_text("Conteúdo 1")
    file2.write_text("Conteúdo 2")
    
    version_manager.create_backup(str(file1), "update", "Teste 1")
    version_manager.create_backup(str(file2), "update", "Teste 2")
    
    files = version_manager.get_all_files_with_history()
    assert len(files) == 2
    assert "file1.txt" in files
    assert "file2.txt" in files


def test_save_after_state(version_manager, sample_file):
    """Testa salvamento do estado pós-operação."""
    version_id = version_manager.create_backup(sample_file, "update", "Teste")
    
    # Modificar arquivo
    Path(sample_file).write_text("Conteúdo modificado")
    
    # Salvar estado pós-operação
    version_manager.save_after_state(sample_file, version_id)
    
    history = version_manager.get_history(sample_file)
    version = history.versions[0]
    
    assert version.after_path is not None
    assert Path(version.after_path).exists()
