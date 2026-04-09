"""Tests for list_charts functionality."""

import pytest
from pathlib import Path
from src.excel_tool import ExcelTool


@pytest.fixture
def excel_tool():
    """Create an ExcelTool instance for testing."""
    return ExcelTool()


@pytest.fixture
def file_with_multiple_charts(tmp_path):
    """Create a sample Excel file with multiple charts for testing."""
    file_path = tmp_path / "test_list_charts.xlsx"
    
    tool = ExcelTool()
    
    # Create file with data
    data = {
        "Sheet1": [
            ["Produto", "Vendas", "Custo"],
            ["A", 100, 50],
            ["B", 200, 80],
            ["C", 150, 60]
        ],
        "Sheet2": [
            ["Mês", "Receita"],
            ["Jan", 1000],
            ["Fev", 1200]
        ]
    }
    tool.create_excel(str(file_path), data)
    
    # Add charts to Sheet1
    chart_config_1 = {
        'type': 'column',
        'title': 'Vendas por Produto',
        'categories': 'A2:A4',
        'values': 'B2:B4',
        'position': 'E2'
    }
    tool.add_chart(str(file_path), 'Sheet1', chart_config_1)
    
    chart_config_2 = {
        'type': 'pie',
        'title': 'Distribuição de Custos',
        'categories': 'A2:A4',
        'values': 'C2:C4',
        'position': 'E15'
    }
    tool.add_chart(str(file_path), 'Sheet1', chart_config_2)
    
    # Add chart to Sheet2
    chart_config_3 = {
        'type': 'line',
        'title': 'Receita Mensal',
        'categories': 'A2:A3',
        'values': 'B2:B3',
        'position': 'D2'
    }
    tool.add_chart(str(file_path), 'Sheet2', chart_config_3)
    
    return file_path


class TestListCharts:
    """Test suite for list_charts method."""
    
    def test_list_charts_in_file_with_no_charts(self, excel_tool, tmp_path):
        """Test listing charts in a file with no charts."""
        file_path = tmp_path / "no_charts.xlsx"
        
        data = {"Sheet1": [["A", "B"], [1, 2]]}
        excel_tool.create_excel(str(file_path), data)
        
        result = excel_tool.list_charts(str(file_path))
        
        assert result['total_count'] == 0
        assert len(result['charts']) == 0
        assert 'Sheet1' in result['sheets_analyzed']
    
    def test_list_all_charts_in_file(self, excel_tool, file_with_multiple_charts):
        """Test listing all charts in a file (all sheets)."""
        result = excel_tool.list_charts(str(file_with_multiple_charts))
        
        assert result['total_count'] == 3
        assert len(result['charts']) == 3
        assert len(result['sheets_analyzed']) == 2
        
        # Verify chart information
        chart_titles = [chart['title'] for chart in result['charts']]
        assert 'Vendas por Produto' in chart_titles
        assert 'Distribuição de Custos' in chart_titles
        assert 'Receita Mensal' in chart_titles
    
    def test_list_charts_in_specific_sheet(self, excel_tool, file_with_multiple_charts):
        """Test listing charts in a specific sheet."""
        result = excel_tool.list_charts(str(file_with_multiple_charts), 'Sheet1')
        
        assert result['total_count'] == 2
        assert len(result['charts']) == 2
        assert result['sheets_analyzed'] == ['Sheet1']
        
        # All charts should be from Sheet1
        for chart in result['charts']:
            assert chart['sheet'] == 'Sheet1'
    
    def test_list_charts_includes_chart_details(self, excel_tool, file_with_multiple_charts):
        """Test that list_charts includes all chart details."""
        result = excel_tool.list_charts(str(file_with_multiple_charts), 'Sheet1')
        
        chart = result['charts'][0]
        
        # Verify all required fields are present
        assert 'sheet' in chart
        assert 'title' in chart
        assert 'type' in chart
        assert 'position' in chart
        assert 'index' in chart
        
        # Verify field types
        assert isinstance(chart['sheet'], str)
        assert isinstance(chart['title'], str)
        assert isinstance(chart['type'], str)
        assert isinstance(chart['position'], str)
        assert isinstance(chart['index'], int)
    
    def test_list_charts_shows_correct_types(self, excel_tool, file_with_multiple_charts):
        """Test that chart types are correctly identified."""
        result = excel_tool.list_charts(str(file_with_multiple_charts))
        
        chart_types = [chart['type'] for chart in result['charts']]
        
        assert 'BarChart' in chart_types  # column charts are BarChart type
        assert 'PieChart' in chart_types
        assert 'LineChart' in chart_types
    
    def test_list_charts_shows_correct_positions(self, excel_tool, file_with_multiple_charts):
        """Test that chart positions are correctly reported."""
        result = excel_tool.list_charts(str(file_with_multiple_charts), 'Sheet1')
        
        positions = [chart['position'] for chart in result['charts']]
        
        assert 'E2' in positions
        assert 'E15' in positions
    
    def test_list_charts_with_untitled_chart(self, excel_tool, tmp_path):
        """Test listing charts when a chart has no title."""
        file_path = tmp_path / "untitled_chart.xlsx"
        
        data = {"Sheet1": [["A", "B"], [1, 2], [3, 4]]}
        excel_tool.create_excel(str(file_path), data)
        
        # Add chart without title
        chart_config = {
            'type': 'column',
            'categories': 'A2:A3',
            'values': 'B2:B3',
            'position': 'D2'
        }
        excel_tool.add_chart(str(file_path), 'Sheet1', chart_config)
        
        result = excel_tool.list_charts(str(file_path))
        
        assert result['total_count'] == 1
        assert result['charts'][0]['title'] == 'Untitled'
    
    def test_list_charts_with_invalid_sheet_name(self, excel_tool, file_with_multiple_charts):
        """Test that listing charts with invalid sheet name raises ValueError."""
        with pytest.raises(ValueError, match="Sheet 'NonExistent' not found"):
            excel_tool.list_charts(str(file_with_multiple_charts), 'NonExistent')
    
    def test_list_charts_with_nonexistent_file(self, excel_tool):
        """Test that listing charts with nonexistent file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            excel_tool.list_charts('nonexistent.xlsx')
    
    def test_list_charts_index_is_correct(self, excel_tool, file_with_multiple_charts):
        """Test that chart indices are correctly assigned."""
        result = excel_tool.list_charts(str(file_with_multiple_charts), 'Sheet1')
        
        indices = [chart['index'] for chart in result['charts']]
        
        # Should have indices 0 and 1
        assert 0 in indices
        assert 1 in indices
        assert len(set(indices)) == 2  # All indices should be unique
    
    def test_list_charts_return_structure(self, excel_tool, file_with_multiple_charts):
        """Test that list_charts returns the correct structure."""
        result = excel_tool.list_charts(str(file_with_multiple_charts))
        
        # Verify top-level structure
        assert 'charts' in result
        assert 'total_count' in result
        assert 'sheets_analyzed' in result
        
        # Verify types
        assert isinstance(result['charts'], list)
        assert isinstance(result['total_count'], int)
        assert isinstance(result['sheets_analyzed'], list)
        
        # Verify consistency
        assert result['total_count'] == len(result['charts'])
    
    def test_list_charts_empty_sheet(self, excel_tool, tmp_path):
        """Test listing charts in a sheet with no charts."""
        file_path = tmp_path / "empty_sheet.xlsx"
        
        data = {
            "Sheet1": [["A", "B"], [1, 2]],
            "Sheet2": [["C", "D"], [3, 4]]
        }
        excel_tool.create_excel(str(file_path), data)
        
        # Add chart only to Sheet1
        chart_config = {
            'type': 'column',
            'title': 'Test Chart',
            'categories': 'A2:A2',
            'values': 'B2:B2',
            'position': 'D2'
        }
        excel_tool.add_chart(str(file_path), 'Sheet1', chart_config)
        
        # List charts in Sheet2 (no charts)
        result = excel_tool.list_charts(str(file_path), 'Sheet2')
        
        assert result['total_count'] == 0
        assert len(result['charts']) == 0
        assert result['sheets_analyzed'] == ['Sheet2']
