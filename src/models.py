"""Data models for Gemini Office Agent.

This module contains dataclasses representing the core data structures
used throughout the application for file information, Office document data,
and agent communication.
"""

from dataclasses import dataclass
from typing import Any, Dict, List, Optional


@dataclass
class FileInfo:
    """Represents information about an Office file discovered by the scanner.
    
    Attributes:
        path: Complete file path
        name: File name
        extension: File extension (.xlsx, .docx, .pptx)
        size: File size in bytes
        modified_time: Timestamp of last modification
    """
    path: str
    name: str
    extension: str
    size: int
    modified_time: float


@dataclass
class ExcelData:
    """Represents data from an Excel file.
    
    Attributes:
        sheets: Dictionary mapping sheet names to their data (list of rows)
        metadata: Additional metadata (author, creation date, etc.)
    """
    sheets: Dict[str, List[List[Any]]]
    metadata: Dict[str, Any]


@dataclass
class WordData:
    """Represents data from a Word document.
    
    Attributes:
        paragraphs: List of paragraph texts
        tables: List of tables, each table is a list of rows, each row is a list of cells
        metadata: Additional metadata
    """
    paragraphs: List[str]
    tables: List[List[List[str]]]
    metadata: Dict[str, Any]


@dataclass
class SlideData:
    """Represents data from a single PowerPoint slide.
    
    Attributes:
        index: Slide index (0-based)
        title: Slide title
        content: List of text content items
        notes: Presenter notes
    """
    index: int
    title: str
    content: List[str]
    notes: str


@dataclass
class PowerPointData:
    """Represents data from a PowerPoint presentation.
    
    Attributes:
        slides: List of slide data
        metadata: Additional metadata
    """
    slides: List[SlideData]
    metadata: Dict[str, Any]


@dataclass
class AgentResponse:
    """Represents the response from the Agent after processing a request.
    
    Attributes:
        success: Whether the operation was successful
        message: Descriptive message about the operation
        files_modified: List of file paths that were modified
        error: Error message if operation failed, None otherwise
        batch_result: Detailed results for batch operations, None for single operations
        is_chat: True when the response is a conversational reply, not a file action
    """
    success: bool
    message: str
    files_modified: List[str]
    error: Optional[str] = None
    batch_result: Optional['BatchResult'] = None
    is_chat: bool = False


@dataclass
class ActionResult:
    """Represents the result of a single action in a batch operation.
    
    Attributes:
        action_index: Index of the action (0-based)
        tool: Tool used (excel, word, powerpoint)
        operation: Operation performed (create, update, add, etc.)
        target_file: File that was targeted
        success: Whether the action succeeded
        error: Error message if action failed, None otherwise
    """
    action_index: int
    tool: str
    operation: str
    target_file: str
    success: bool
    error: Optional[str] = None


@dataclass
class BatchResult:
    """Represents the results of a batch operation with multiple actions.
    
    Attributes:
        total_actions: Total number of actions attempted
        successful_actions: Number of actions that succeeded
        failed_actions: Number of actions that failed
        action_results: Detailed results for each action
        overall_success: True if at least one action succeeded
    """
    total_actions: int
    successful_actions: int
    failed_actions: int
    action_results: List[ActionResult]
    overall_success: bool


@dataclass
class GeminiRequest:
    """Represents a request to the Gemini API.
    
    Attributes:
        prompt: Complete prompt to send to Gemini
        user_intent: Original user intent/command
        context_files: List of file paths included in the context
        timestamp: Request timestamp
    """
    prompt: str
    user_intent: str
    context_files: List[str]
    timestamp: float
