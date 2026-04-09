"""Unit tests for data models."""

import time
from src.models import (
    FileInfo,
    ExcelData,
    WordData,
    SlideData,
    PowerPointData,
    AgentResponse,
    GeminiRequest,
)


def test_file_info_creation():
    """Test FileInfo dataclass creation."""
    file_info = FileInfo(
        path="/path/to/file.xlsx",
        name="file.xlsx",
        extension=".xlsx",
        size=1024,
        modified_time=time.time(),
    )
    assert file_info.path == "/path/to/file.xlsx"
    assert file_info.name == "file.xlsx"
    assert file_info.extension == ".xlsx"
    assert file_info.size == 1024
    assert isinstance(file_info.modified_time, float)


def test_excel_data_creation():
    """Test ExcelData dataclass creation."""
    excel_data = ExcelData(
        sheets={"Sheet1": [[1, 2, 3], [4, 5, 6]]},
        metadata={"author": "Test User"},
    )
    assert "Sheet1" in excel_data.sheets
    assert excel_data.sheets["Sheet1"] == [[1, 2, 3], [4, 5, 6]]
    assert excel_data.metadata["author"] == "Test User"


def test_word_data_creation():
    """Test WordData dataclass creation."""
    word_data = WordData(
        paragraphs=["Paragraph 1", "Paragraph 2"],
        tables=[[["Cell 1", "Cell 2"]]],
        metadata={"title": "Test Document"},
    )
    assert len(word_data.paragraphs) == 2
    assert word_data.paragraphs[0] == "Paragraph 1"
    assert len(word_data.tables) == 1
    assert word_data.metadata["title"] == "Test Document"


def test_slide_data_creation():
    """Test SlideData dataclass creation."""
    slide_data = SlideData(
        index=0,
        title="Test Slide",
        content=["Content 1", "Content 2"],
        notes="Speaker notes",
    )
    assert slide_data.index == 0
    assert slide_data.title == "Test Slide"
    assert len(slide_data.content) == 2
    assert slide_data.notes == "Speaker notes"


def test_powerpoint_data_creation():
    """Test PowerPointData dataclass creation."""
    slide1 = SlideData(
        index=0,
        title="Slide 1",
        content=["Content"],
        notes="Notes",
    )
    slide2 = SlideData(
        index=1,
        title="Slide 2",
        content=["More content"],
        notes="More notes",
    )
    ppt_data = PowerPointData(
        slides=[slide1, slide2],
        metadata={"author": "Test User"},
    )
    assert len(ppt_data.slides) == 2
    assert ppt_data.slides[0].title == "Slide 1"
    assert ppt_data.slides[1].title == "Slide 2"
    assert ppt_data.metadata["author"] == "Test User"


def test_agent_response_success():
    """Test AgentResponse for successful operation."""
    response = AgentResponse(
        success=True,
        message="Operation completed successfully",
        files_modified=["file1.xlsx", "file2.docx"],
        error=None,
    )
    assert response.success is True
    assert response.message == "Operation completed successfully"
    assert len(response.files_modified) == 2
    assert response.error is None


def test_agent_response_failure():
    """Test AgentResponse for failed operation."""
    response = AgentResponse(
        success=False,
        message="Operation failed",
        files_modified=[],
        error="File not found",
    )
    assert response.success is False
    assert response.message == "Operation failed"
    assert len(response.files_modified) == 0
    assert response.error == "File not found"


def test_gemini_request_creation():
    """Test GeminiRequest dataclass creation."""
    timestamp = time.time()
    request = GeminiRequest(
        prompt="Complete prompt with context",
        user_intent="Create a new Excel file",
        context_files=["file1.xlsx", "file2.docx"],
        timestamp=timestamp,
    )
    assert request.prompt == "Complete prompt with context"
    assert request.user_intent == "Create a new Excel file"
    assert len(request.context_files) == 2
    assert request.timestamp == timestamp


def test_empty_collections():
    """Test models with empty collections."""
    excel_data = ExcelData(sheets={}, metadata={})
    assert len(excel_data.sheets) == 0
    assert len(excel_data.metadata) == 0
    
    word_data = WordData(paragraphs=[], tables=[], metadata={})
    assert len(word_data.paragraphs) == 0
    assert len(word_data.tables) == 0
    
    ppt_data = PowerPointData(slides=[], metadata={})
    assert len(ppt_data.slides) == 0
    
    response = AgentResponse(
        success=True,
        message="No files modified",
        files_modified=[],
    )
    assert len(response.files_modified) == 0
