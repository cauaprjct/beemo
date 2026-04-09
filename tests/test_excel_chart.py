"""Tests for Excel chart functionality."""

import pytest
import tempfile
import os
from pathlib import Path
from src.excel_tool import ExcelTool
from src.exceptions import ValidationError


class TestExcelChart:
    """Test chart operations in ExcelTool."""
    
    def setup_method(self):
        """Setup test fixtures."""
        self.tool = ExcelTool()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test_chart.xlsx")
        
        # Create a test file with sample data
        self.tool.create_excel(self.test_file, {
            "Sales": [
                ["Product", "Sales", "Target"],
                ["Product A", 100, 120],
                ["Product B", 150, 140],
                ["Product C", 200, 180],
                ["Product D", 80, 100],
                ["Product E", 120, 130]
            ]
        })
    
    def teardown_method(self):
        """Cleanup test files."""
        # Note: On Windows, files might be locked. Ignore errors.
        try:
            if os.path.exists(self.test_file):
                os.remove(self.test_file)
            os.rmdir(self.temp_dir)
        except:
            pass
    
    def test_add_column_chart(self):
        """Test adding a column chart."""
        chart_config = {
            "type": "column",
            "title": "Sales by Product",
            "categories": "A2:A6",
            "values": "B2:B6",
            "position": "E2"
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
        
        # Verify file still exists and is valid
        data = self.tool.read_excel(self.test_file)
        assert "Sales" in data['sheets']
    
    def test_add_pie_chart(self):
        """Test adding a pie chart."""
        chart_config = {
            "type": "pie",
            "title": "Sales Distribution",
            "data_range": "A1:B6",
            "position": "E2"
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
        
        # Verify file still exists and is valid
        data = self.tool.read_excel(self.test_file)
        assert "Sales" in data['sheets']
    
    def test_add_line_chart_with_axis_titles(self):
        """Test adding a line chart with axis titles."""
        chart_config = {
            "type": "line",
            "title": "Sales Trend",
            "categories": "A2:A6",
            "values": "B2:B6",
            "x_axis_title": "Product",
            "y_axis_title": "Sales",
            "position": "E2",
            "width": 20,
            "height": 12
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
        
        # Verify file still exists and is valid
        data = self.tool.read_excel(self.test_file)
        assert "Sales" in data['sheets']
    
    def test_add_bar_chart(self):
        """Test adding a horizontal bar chart."""
        chart_config = {
            "type": "bar",
            "title": "Sales Comparison",
            "categories": "A2:A6",
            "values": "B2:B6",
            "position": "E2"
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
    
    def test_add_area_chart(self):
        """Test adding an area chart."""
        chart_config = {
            "type": "area",
            "title": "Sales Area",
            "categories": "A2:A6",
            "values": "B2:B6",
            "position": "E2"
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
    
    def test_add_scatter_chart(self):
        """Test adding a scatter chart."""
        chart_config = {
            "type": "scatter",
            "title": "Sales vs Target",
            "categories": "B2:B6",
            "values": "C2:C6",
            "position": "E2"
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)
    
    def test_add_chart_invalid_type(self):
        """Test that invalid chart type raises ValueError."""
        chart_config = {
            "type": "invalid_type",
            "data_range": "A1:B6"
        }
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.add_chart(self.test_file, "Sales", chart_config)
        
        assert "invalid chart type" in str(exc_info.value).lower()
    
    def test_add_chart_invalid_sheet(self):
        """Test that invalid sheet name raises ValueError."""
        chart_config = {
            "type": "column",
            "data_range": "A1:B6"
        }
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.add_chart(self.test_file, "NonExistentSheet", chart_config)
        
        assert "not found" in str(exc_info.value).lower()
    
    def test_add_chart_missing_data(self):
        """Test that missing data specification raises ValueError."""
        chart_config = {
            "type": "column",
            "title": "Test Chart"
            # Missing data_range or values
        }
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.add_chart(self.test_file, "Sales", chart_config)
        
        assert "data_range" in str(exc_info.value).lower() or "values" in str(exc_info.value).lower()
    
    def test_add_chart_nonexistent_file(self):
        """Test that nonexistent file raises FileNotFoundError."""
        chart_config = {
            "type": "column",
            "data_range": "A1:B6"
        }
        
        with pytest.raises(FileNotFoundError):
            self.tool.add_chart("nonexistent.xlsx", "Sales", chart_config)
    
    def test_add_chart_with_custom_style(self):
        """Test adding a chart with custom style."""
        chart_config = {
            "type": "column",
            "title": "Styled Chart",
            "categories": "A2:A6",
            "values": "B2:B6",
            "position": "E2",
            "style": 15
        }
        
        # Should not raise any exception
        self.tool.add_chart(self.test_file, "Sales", chart_config)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
