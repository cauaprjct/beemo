"""Unit tests for prompt_templates module.

Tests the prompt template generation and context building functionality.
"""

import pytest
from src.prompt_templates import PromptTemplates


class TestPromptTemplates:
    """Test suite for PromptTemplates class."""
    
    def test_get_system_prompt_returns_string(self):
        """Test that system prompt returns a non-empty string."""
        # Act
        prompt = PromptTemplates.get_system_prompt()
        
        # Assert
        assert isinstance(prompt, str)
        assert len(prompt) > 0
    
    def test_get_system_prompt_contains_capabilities(self):
        """Test that system prompt includes information about available capabilities."""
        # Act
        prompt = PromptTemplates.get_system_prompt()
        
        # Assert
        assert "Excel" in prompt or "excel" in prompt
        assert "Word" in prompt or "word" in prompt
        assert "PowerPoint" in prompt or "powerpoint" in prompt
    
    def test_get_system_prompt_contains_response_format(self):
        """Test that system prompt includes JSON response format instructions."""
        # Act
        prompt = PromptTemplates.get_system_prompt()
        
        # Assert
        assert "JSON" in prompt or "json" in prompt
        assert "actions" in prompt
        assert "tool" in prompt
        assert "operation" in prompt
    
    def test_get_system_prompt_contains_examples(self):
        """Test that system prompt includes example responses."""
        # Act
        prompt = PromptTemplates.get_system_prompt()
        
        # Assert
        assert "Example" in prompt or "example" in prompt
        assert "explanation" in prompt
    
    def test_build_context_prompt_with_no_files(self):
        """Test building context prompt without any file contexts."""
        # Arrange
        user_prompt = "Create a new Excel file"
        file_contexts = []
        
        # Act
        result = PromptTemplates.build_context_prompt(user_prompt, file_contexts)
        
        # Assert
        assert isinstance(result, str)
        assert user_prompt in result
        assert "USER REQUEST" in result
    
    def test_build_context_prompt_with_single_file(self):
        """Test building context prompt with one file context."""
        # Arrange
        user_prompt = "Read the sales data"
        file_contexts = [
            {
                'path': 'data/sales.xlsx',
                'type': 'excel',
                'content': 'Sheet1: Row 1: [Product, Sales]'
            }
        ]
        
        # Act
        result = PromptTemplates.build_context_prompt(user_prompt, file_contexts)
        
        # Assert
        assert user_prompt in result
        assert 'data/sales.xlsx' in result
        assert 'excel' in result
        assert 'Sheet1' in result
    
    def test_build_context_prompt_with_multiple_files(self):
        """Test building context prompt with multiple file contexts."""
        # Arrange
        user_prompt = "Create a report from sales and budget data"
        file_contexts = [
            {
                'path': 'data/sales.xlsx',
                'type': 'excel',
                'content': 'Sales data'
            },
            {
                'path': 'data/budget.xlsx',
                'type': 'excel',
                'content': 'Budget data'
            }
        ]
        
        # Act
        result = PromptTemplates.build_context_prompt(user_prompt, file_contexts)
        
        # Assert
        assert 'data/sales.xlsx' in result
        assert 'data/budget.xlsx' in result
        assert 'Sales data' in result
        assert 'Budget data' in result
    
    def test_format_file_content_excel(self):
        """Test formatting Excel file content for prompt context."""
        # Arrange
        file_path = 'test.xlsx'
        file_type = 'excel'
        content = {
            'sheets': {
                'Sheet1': [
                    ['Name', 'Age'],
                    ['Alice', 30],
                    ['Bob', 25]
                ]
            }
        }
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert 'Sheet1' in result['content']
        assert 'Alice' in result['content']
    
    def test_format_file_content_word(self):
        """Test formatting Word file content for prompt context."""
        # Arrange
        file_path = 'test.docx'
        file_type = 'word'
        content = {
            'paragraphs': [
                'First paragraph',
                'Second paragraph'
            ],
            'tables': []
        }
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert 'First paragraph' in result['content']
        assert 'Second paragraph' in result['content']
    
    def test_format_file_content_word_with_tables(self):
        """Test formatting Word file content with tables."""
        # Arrange
        file_path = 'test.docx'
        file_type = 'word'
        content = {
            'paragraphs': ['Document text'],
            'tables': [
                [['Header1', 'Header2'], ['Data1', 'Data2']]
            ]
        }
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert 'Tables' in result['content'] or 'tables' in result['content']
        assert 'Header1' in result['content']
    
    def test_format_file_content_powerpoint(self):
        """Test formatting PowerPoint file content for prompt context."""
        # Arrange
        file_path = 'test.pptx'
        file_type = 'powerpoint'
        content = {
            'slides': [
                {
                    'index': 0,
                    'title': 'Introduction',
                    'content': ['Point 1', 'Point 2']
                },
                {
                    'index': 1,
                    'title': 'Conclusion',
                    'content': ['Summary']
                }
            ]
        }
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert 'Introduction' in result['content']
        assert 'Conclusion' in result['content']
        assert 'Point 1' in result['content']
    
    def test_format_file_content_pdf(self):
        """Test formatting PDF file content for prompt context."""
        # Arrange
        file_path = 'test.pdf'
        file_type = 'pdf'
        content = {
            'pages': ['Page 1 content', 'Page 2 content'],
            'num_pages': 2,
            'metadata': {
                'title': 'Test Document',
                'author': 'Test Author'
            }
        }
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert 'Total pages: 2' in result['content']
        assert 'Page 1 content' in result['content']
        assert 'Page 2 content' in result['content']
        assert 'title: Test Document' in result['content']
        assert 'author: Test Author' in result['content']
    
    def test_format_file_content_unknown_type(self):
        """Test formatting content for unknown file type."""
        # Arrange
        file_path = 'test.txt'
        file_type = 'unknown'
        content = 'Some text content'
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert 'Some text content' in result['content']
    
    def test_build_context_prompt_includes_system_prompt(self):
        """Test that built context includes the system prompt."""
        # Arrange
        user_prompt = "Test request"
        file_contexts = []
        
        # Act
        result = PromptTemplates.build_context_prompt(user_prompt, file_contexts)
        system_prompt = PromptTemplates.get_system_prompt()
        
        # Assert
        assert system_prompt in result
    
    def test_format_file_content_empty_excel(self):
        """Test formatting empty Excel content."""
        # Arrange
        file_path = 'empty.xlsx'
        file_type = 'excel'
        content = {'sheets': {}}
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert isinstance(result['content'], str)
    
    def test_format_file_content_empty_word(self):
        """Test formatting empty Word content."""
        # Arrange
        file_path = 'empty.docx'
        file_type = 'word'
        content = {'paragraphs': [], 'tables': []}
        
        # Act
        result = PromptTemplates.format_file_content(file_path, file_type, content)
        
        # Assert
        assert result['path'] == file_path
        assert result['type'] == file_type
        assert isinstance(result['content'], str)
