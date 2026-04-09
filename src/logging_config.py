"""Configuração de logging para o Gemini Office Agent."""

import logging
import traceback
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional


def setup_logging(
    log_file: Optional[str] = None,
    log_level: int = logging.INFO,
    console_output: bool = True
) -> None:
    """Configura o sistema de logging do aplicativo.

    Configura dois handlers:
    - Console (StreamHandler) com nível INFO
    - Arquivo (RotatingFileHandler) com nível DEBUG em logs/agent.log

    Args:
        log_file: Caminho do arquivo de log. Se None, usa 'logs/agent.log'
        log_level: Nível de log para o console (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        console_output: Se True, também exibe logs no console
    """
    if log_file is None:
        log_file = "logs/agent.log"

    # Cria o diretório de logs se não existir
    log_path = Path(log_file)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    # Formato do log com timestamp
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    formatter = logging.Formatter(log_format, datefmt=date_format)

    # Configura o logger raiz
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Remove handlers existentes para evitar duplicação
    root_logger.handlers.clear()

    # Handler para arquivo com rotação (máx 5 MB, 3 backups)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    # Handler para console (opcional)
    if console_output:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(log_level)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)

    logging.info("Sistema de logging inicializado")
    logging.info(f"Arquivo de log: {log_file}")
    logging.info(f"Nível de log (console): {logging.getLevelName(log_level)}")


def get_logger(name: str) -> logging.Logger:
    """Retorna um logger configurado para o módulo especificado.

    Args:
        name: Nome do módulo (geralmente __name__)

    Returns:
        Logger configurado
    """
    return logging.getLogger(name)


def log_operation_start(logger: logging.Logger, operation: str) -> None:
    """Registra o início de uma operação principal.

    Args:
        logger: Logger do módulo chamador
        operation: Nome/descrição da operação que está iniciando
    """
    logger.info(f"[INÍCIO] {operation}")


def log_operation_end(logger: logging.Logger, operation: str) -> None:
    """Registra o fim de uma operação principal.

    Args:
        logger: Logger do módulo chamador
        operation: Nome/descrição da operação que foi concluída
    """
    logger.info(f"[FIM] {operation}")


def log_error_with_traceback(
    logger: logging.Logger, error: Exception, context: str = ""
) -> None:
    """Registra um erro com stack trace completo.

    Args:
        logger: Logger do módulo chamador
        error: Exceção capturada
        context: Contexto adicional sobre onde/quando o erro ocorreu
    """
    context_msg = f" | Contexto: {context}" if context else ""
    logger.error(
        f"Erro: {type(error).__name__}: {error}{context_msg}\n"
        f"{traceback.format_exc()}"
    )


def log_file_access_error(
    logger: logging.Logger, file_path: str, reason: str
) -> None:
    """Registra erro de acesso a arquivo com caminho e razão da falha.

    Args:
        logger: Logger do módulo chamador
        file_path: Caminho do arquivo que não pôde ser acessado
        reason: Razão da falha de acesso
    """
    logger.error(f"Falha ao acessar arquivo '{file_path}': {reason}")
