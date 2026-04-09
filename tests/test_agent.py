"""Unit tests for Agent class.

Tests the main workflow coordination logic of the Agent, including
file discovery, content reading, prompt building, and action execution.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch, call
from pathlib import Path

from src.agent import Agent
from src.models import AgentResponse
from src.exceptions import ValidationError, CorruptedFileError
from config import Config


@pytest.fixture
def mock_config():
    """Create a mock Config object."""
    config = Mock(spec=Config)
    config.api_key = "test_api_key"
    config.model_name = "gemini-2.5-flash-lite"
    config.fallback_models = ["gemini-2.5-flash", "gemini-2.5-pro"]
    config.root_path = "/test/root"
    return config


@pytest.fixture
def agent(mock_config):
    """Create an Agent instance with mocked dependencies."""
    with patch('src.agent.FileScanner'), \
         patch('src.agent.GeminiClient'), \
         patch('src.agent.ExcelTool'), \
         patch('src.agent.WordTool'), \
         patch('src.agent.PowerPointTool'), \
         patch('src.agent.PdfTool'), \
         patch('src.agent.SecurityValidator'):
        
        agent = Agent(mock_config)
        return agent


class TestAgentInitialization:
    """Tests for Agent initialization."""
    
    def test_agent_initializes_with_config(self, mock_config):
        """Test that Agent initializes all components correctly."""
        with patch('src.agent.FileScanner') as mock_scanner, \
             patch('src.agent.GeminiClient') as mock_client, \
             patch('src.agent.ExcelTool') as mock_excel, \
             patch('src.agent.WordTool') as mock_word, \
             patch('src.agent.PowerPointTool') as mock_ppt, \
             patch('src.agent.PdfTool') as mock_pdf, \
             patch('src.agent.SecurityValidator') as mock_validator:
            
            agent = Agent(mock_config)
            
            # Verify all components were initialized
            mock_scanner.assert_called_once_with(mock_config.root_path)
            mock_client.assert_called_once_with(
                mock_config.api_key, mock_config.model_name, None,
                fallback_models=mock_config.fallback_models
            )
            mock_excel.assert_called_once()
            mock_word.assert_called_once()
            mock_ppt.assert_called_once()
            mock_validator.assert_called_once_with(mock_config.root_path)
            
            assert agent.config == mock_config


class TestDiscoverFiles:
    """Tests for _discover_files method."""
    
    def test_discover_files_calls_scanner(self, agent):
        """Test that _discover_files calls FileScanner.scan_office_files."""
        agent.file_scanner.scan_office_files.return_value = [
            "/test/file1.xlsx",
            "/test/file2.docx"
        ]
        
        result = agent._discover_files()
        
        agent.file_scanner.scan_office_files.assert_called_once()
        assert result == ["/test/file1.xlsx", "/test/file2.docx"]
    
    def test_discover_files_returns_empty_list(self, agent):
        """Test that _discover_files returns empty list when no files found."""
        agent.file_scanner.scan_office_files.return_value = []
        
        result = agent._discover_files()
        
        assert result == []


class TestFilterRelevantFiles:
    """Tests for _filter_relevant_files method."""
    
    def test_filter_by_exact_filename(self, agent):
        """Test filtering files by exact filename mention."""
        available_files = [
            "/test/sales.xlsx",
            "/test/report.docx",
            "/test/presentation.pptx"
        ]
        user_prompt = "Read the sales.xlsx file"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/sales.xlsx" in result
        assert len(result) == 1
    
    def test_filter_by_file_stem(self, agent):
        """Test filtering files by filename without extension."""
        available_files = [
            "/test/budget.xlsx",
            "/test/notes.docx"
        ]
        user_prompt = "Update the budget file"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/budget.xlsx" in result
    
    def test_filter_by_file_type_excel(self, agent):
        """Test filtering Excel files by type keywords."""
        available_files = [
            "/test/data.xlsx",
            "/test/doc.docx"
        ]
        user_prompt = "Show me the Excel spreadsheet"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/data.xlsx" in result
        assert "/test/doc.docx" not in result
    
    def test_filter_by_file_type_word(self, agent):
        """Test filtering Word files by type keywords."""
        available_files = [
            "/test/data.xlsx",
            "/test/report.docx"
        ]
        user_prompt = "Read the Word document"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/report.docx" in result
        assert "/test/data.xlsx" not in result
    
    def test_filter_by_file_type_powerpoint(self, agent):
        """Test filtering PowerPoint files by type keywords."""
        available_files = [
            "/test/data.xlsx",
            "/test/slides.pptx"
        ]
        user_prompt = "Show me the PowerPoint presentation"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/slides.pptx" in result
        assert "/test/data.xlsx" not in result
    
    def test_filter_returns_all_when_no_match(self, agent):
        """Test that all files are returned when no specific match found."""
        available_files = [
            "/test/file1.xlsx",
            "/test/file2.docx"
        ]
        user_prompt = "Do something"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert result == available_files
    
    def test_filter_case_insensitive(self, agent):
        """Test that filtering is case-insensitive."""
        available_files = ["/test/SALES.xlsx"]
        user_prompt = "read sales.xlsx"
        
        result = agent._filter_relevant_files(available_files, user_prompt)
        
        assert "/test/SALES.xlsx" in result


class TestReadFileContent:
    """Tests for _read_file_content method."""
    
    def test_read_excel_file(self, agent):
        """Test reading Excel file content."""
        agent.excel_tool.read_excel.return_value = {
            'sheets': {'Sheet1': [[1, 2, 3]]},
            'metadata': {}
        }
        
        result = agent._read_file_content("/test/file.xlsx")
        
        agent.excel_tool.read_excel.assert_called_once_with("/test/file.xlsx")
        assert 'sheets' in result
    
    def test_read_word_file(self, agent):
        """Test reading Word file content."""
        agent.word_tool.read_word.return_value = "Paragraph 1\nParagraph 2"
        agent.word_tool.extract_tables.return_value = []
        
        result = agent._read_file_content("/test/file.docx")
        
        agent.word_tool.read_word.assert_called_once_with("/test/file.docx")
        agent.word_tool.extract_tables.assert_called_once_with("/test/file.docx")
        assert 'paragraphs' in result
        assert result['paragraphs'] == ["Paragraph 1", "Paragraph 2"]
    
    def test_read_powerpoint_file(self, agent):
        """Test reading PowerPoint file content."""
        agent.powerpoint_tool.read_powerpoint.return_value = [
            {'index': 0, 'title': 'Slide 1', 'content': ['Content'], 'notes': ''}
        ]
        
        result = agent._read_file_content("/test/file.pptx")
        
        agent.powerpoint_tool.read_powerpoint.assert_called_once_with("/test/file.pptx")
        assert 'slides' in result
    
    def test_read_unsupported_extension(self, agent):
        """Test that unsupported file extension raises ValueError."""
        with pytest.raises(ValueError, match="Extensão de arquivo não suportada"):
            agent._read_file_content("/test/file.txt")


class TestGetFileType:
    """Tests for _get_file_type method."""
    
    def test_get_file_type_excel(self, agent):
        """Test identifying Excel file type."""
        assert agent._get_file_type("/test/file.xlsx") == "excel"
    
    def test_get_file_type_word(self, agent):
        """Test identifying Word file type."""
        assert agent._get_file_type("/test/file.docx") == "word"
    
    def test_get_file_type_powerpoint(self, agent):
        """Test identifying PowerPoint file type."""
        assert agent._get_file_type("/test/file.pptx") == "powerpoint"
    
    def test_get_file_type_pdf(self, agent):
        """Test identifying PDF file type."""
        assert agent._get_file_type("/test/file.pdf") == "pdf"
    
    def test_get_file_type_unknown(self, agent):
        """Test identifying unknown file type."""
        assert agent._get_file_type("/test/file.txt") == "unknown"
    
    def test_get_file_type_case_insensitive(self, agent):
        """Test that file type detection is case-insensitive."""
        assert agent._get_file_type("/test/file.XLSX") == "excel"


class TestBuildContextPrompt:
    """Tests for _build_context_prompt method."""
    
    @patch('src.agent.PromptTemplates')
    def test_build_context_prompt_calls_template(self, mock_templates, agent):
        """Test that _build_context_prompt calls PromptTemplates.build_context_prompt."""
        mock_templates.build_context_prompt.return_value = "Built prompt"
        
        file_contexts = [
            {'path': '/test/file.xlsx', 'type': 'excel', 'content': 'data'}
        ]
        
        result = agent._build_context_prompt("user prompt", file_contexts)
        
        mock_templates.build_context_prompt.assert_called_once_with(
            "user prompt", file_contexts
        )
        assert result == "Built prompt"


class TestExecuteSingleAction:
    """Tests for _execute_single_action method."""
    
    def test_execute_excel_read(self, agent):
        """Test executing Excel read operation."""
        agent._execute_single_action('excel', 'read', '/test/file.xlsx', {})
        
        agent.excel_tool.read_excel.assert_called_once_with('/test/file.xlsx')
    
    def test_execute_excel_create(self, agent):
        """Test executing Excel create operation."""
        data = {'Sheet1': [[1, 2, 3]]}
        agent._execute_single_action('excel', 'create', '/test/file.xlsx', {'data': data})
        
        agent.excel_tool.create_excel.assert_called_once_with('/test/file.xlsx', data)
    
    def test_execute_excel_update(self, agent):
        """Test executing Excel update operation."""
        params = {'sheet': 'Sheet1', 'row': 1, 'col': 1, 'value': 100}
        agent._execute_single_action('excel', 'update', '/test/file.xlsx', params)
        
        agent.excel_tool.update_cell.assert_called_once_with(
            '/test/file.xlsx', 'Sheet1', 1, 1, 100
        )
    
    def test_execute_excel_add(self, agent):
        """Test executing Excel add sheet operation."""
        params = {'sheet_name': 'NewSheet', 'data': [[1, 2]]}
        agent._execute_single_action('excel', 'add', '/test/file.xlsx', params)
        
        agent.excel_tool.add_sheet.assert_called_once_with(
            '/test/file.xlsx', 'NewSheet', [[1, 2]]
        )
    
    def test_execute_word_create(self, agent):
        """Test executing Word create operation."""
        params = {'content': 'Document content'}
        agent._execute_single_action('word', 'create', '/test/file.docx', params)
        
        agent.word_tool.create_word.assert_called_once_with(
            '/test/file.docx', 'Document content'
        )
    
    def test_execute_word_add(self, agent):
        """Test executing Word add paragraph operation."""
        params = {'text': 'New paragraph'}
        agent._execute_single_action('word', 'add', '/test/file.docx', params)
        
        agent.word_tool.add_paragraph.assert_called_once_with(
            '/test/file.docx', 'New paragraph'
        )
    
    def test_execute_powerpoint_create(self, agent):
        """Test executing PowerPoint create operation."""
        slides = [{'title': 'Slide 1', 'content': ['Content']}]
        params = {'slides': slides}
        agent._execute_single_action('powerpoint', 'create', '/test/file.pptx', params)
        
        agent.powerpoint_tool.create_powerpoint.assert_called_once_with(
            '/test/file.pptx', slides
        )
    
    def test_execute_pdf_read(self, agent):
        """Test executing PDF read operation."""
        agent._execute_single_action('pdf', 'read', '/test/file.pdf', {})
        
        agent.pdf_tool.read_pdf.assert_called_once_with('/test/file.pdf')
    
    def test_execute_pdf_create(self, agent):
        """Test executing PDF create operation."""
        elements = [{'type': 'title', 'text': 'Test'}]
        params = {'elements': elements, 'page_size': 'A4'}
        agent._execute_single_action('pdf', 'create', '/test/file.pdf', params)
        
        agent.pdf_tool.create_pdf.assert_called_once_with(
            '/test/file.pdf', elements, 'A4'
        )
    
    def test_execute_pdf_merge(self, agent):
        """Test executing PDF merge operation."""
        params = {'file_paths': ['file1.pdf', 'file2.pdf']}
        agent._execute_single_action('pdf', 'merge', '/test/output.pdf', params)
        
        # Verifica que merge foi chamado (paths são validados internamente)
        assert agent.pdf_tool.merge_pdfs.called
    
    def test_execute_invalid_tool(self, agent):
        """Test that invalid tool raises ValueError."""
        with pytest.raises(ValueError, match="Ferramenta não reconhecida"):
            agent._execute_single_action('invalid', 'read', '/test/file.xlsx', {})
    
    def test_execute_invalid_operation(self, agent):
        """Test that invalid operation raises ValueError."""
        with pytest.raises(ValueError, match="Operação Excel não reconhecida"):
            agent._execute_single_action('excel', 'invalid', '/test/file.xlsx', {})


class TestExecuteActions:
    """Tests for _execute_actions method."""
    
    @patch('src.agent.ResponseParser')
    def test_execute_actions_success(self, mock_parser, agent):
        """Test successful action execution."""
        # Mock parser response
        mock_parser.parse_response.return_value = {
            'actions': [
                {
                    'tool': 'excel',
                    'operation': 'read',
                    'target_file': '/test/file.xlsx',
                    'parameters': {}
                }
            ],
            'explanation': 'Reading Excel file'
        }
        mock_parser.extract_actions.return_value = [
            {
                'tool': 'excel',
                'operation': 'read',
                'target_file': '/test/file.xlsx',
                'parameters': {}
            }
        ]
        mock_parser.extract_explanation.return_value = 'Reading Excel file'
        
        # Mock security validator
        agent.security_validator.validate_file_path.return_value = Path('/test/file.xlsx')
        
        result = agent._execute_actions('{"actions": [...]}')
        
        assert result.success is True
        assert result.message == 'Reading Excel file'
        # Check that the file path is in modified files (normalize path separators)
        assert any('file.xlsx' in f for f in result.files_modified)
    
    @patch('src.agent.ResponseParser')
    def test_execute_actions_validation_error(self, mock_parser, agent):
        """Test action execution with validation error."""
        mock_parser.parse_response.side_effect = ValidationError("Invalid response")
        
        result = agent._execute_actions('invalid response')
        
        assert result.success is False
        assert 'validação' in result.message.lower()
        assert result.error is not None
    
    @patch('src.agent.ResponseParser')
    def test_execute_actions_multiple_actions(self, mock_parser, agent):
        """Test executing multiple actions."""
        mock_parser.parse_response.return_value = {
            'actions': [
                {'tool': 'excel', 'operation': 'read', 'target_file': '/test/file1.xlsx', 'parameters': {}},
                {'tool': 'word', 'operation': 'read', 'target_file': '/test/file2.docx', 'parameters': {}}
            ],
            'explanation': 'Reading files'
        }
        mock_parser.extract_actions.return_value = [
            {'tool': 'excel', 'operation': 'read', 'target_file': '/test/file1.xlsx', 'parameters': {}},
            {'tool': 'word', 'operation': 'read', 'target_file': '/test/file2.docx', 'parameters': {}}
        ]
        mock_parser.extract_explanation.return_value = 'Reading files'
        
        agent.security_validator.validate_file_path.side_effect = [
            Path('/test/file1.xlsx'),
            Path('/test/file2.docx')
        ]
        
        result = agent._execute_actions('{"actions": [...]}')
        
        assert result.success is True
        assert len(result.files_modified) == 2


class TestProcessUserRequest:
    """Tests for process_user_request method (integration)."""
    
    @patch('src.agent.PromptTemplates')
    @patch('src.agent.ResponseParser')
    def test_process_user_request_success(self, mock_parser, mock_templates, agent):
        """Test successful end-to-end user request processing."""
        # Setup mocks
        agent.file_scanner.scan_office_files.return_value = ['/test/file.xlsx']
        agent.excel_tool.read_excel.return_value = {'sheets': {}, 'metadata': {}}
        
        mock_templates.format_file_content.return_value = {
            'path': '/test/file.xlsx',
            'type': 'excel',
            'content': 'data'
        }
        mock_templates.build_context_prompt.return_value = 'context prompt'
        
        agent.gemini_client.generate_response.return_value = '{"actions": [...]}'
        
        mock_parser.parse_response.return_value = {
            'actions': [
                {'tool': 'excel', 'operation': 'read', 'target_file': '/test/file.xlsx', 'parameters': {}}
            ],
            'explanation': 'Success'
        }
        mock_parser.extract_actions.return_value = [
            {'tool': 'excel', 'operation': 'read', 'target_file': '/test/file.xlsx', 'parameters': {}}
        ]
        mock_parser.extract_explanation.return_value = 'Success'
        
        agent.security_validator.validate_file_path.return_value = Path('/test/file.xlsx')
        
        result = agent.process_user_request("Read the Excel file")
        
        assert isinstance(result, AgentResponse)
        assert result.success is True
    
    def test_process_user_request_error_handling(self, agent):
        """Test error handling in process_user_request."""
        agent.file_scanner.scan_office_files.side_effect = Exception("Scanner error")
        
        result = agent.process_user_request("Read files")
        
        assert result.success is False
        assert result.error is not None
