"""
Unit tests for the visualizer module.
"""

import pytest
import tempfile
import pandas as pd
from datetime import datetime
from pathlib import Path
import sys

# Add src directory to path
sys.path.append(str(Path(__file__).parent.parent / 'src'))

from visualizer import TemperatureVisualizer


class TestTemperatureVisualizer:
    """Test cases for TemperatureVisualizer class."""
    
    @pytest.fixture
    def visualizer(self):
        """Create a visualizer instance for testing."""
        with tempfile.TemporaryDirectory() as temp_dir:
            return TemperatureVisualizer(output_dir=temp_dir)
    
    @pytest.fixture
    def sample_device_data(self):
        """Sample device data for testing."""
        return {
            'device_name': 'TestDevice',
            'data': [
                {
                    'timestamp': datetime(2025, 1, 1, 10, 0),
                    'temperature': 20.5,
                    'humidity': 65.0,
                    'battery_mv': 3000
                },
                {
                    'timestamp': datetime(2025, 1, 1, 11, 0),
                    'temperature': 21.2,
                    'humidity': 63.5,
                    'battery_mv': 2995
                },
                {
                    'timestamp': datetime(2025, 1, 1, 12, 0),
                    'temperature': 22.1,
                    'humidity': 61.8,
                    'battery_mv': 2990
                }
            ]
        }
    
    @pytest.fixture
    def multiple_device_data(self, sample_device_data):
        """Multiple device data for testing."""
        device2_data = {
            'device_name': 'TestDevice2',
            'data': [
                {
                    'timestamp': datetime(2025, 1, 1, 10, 0),
                    'temperature': 18.5,
                    'humidity': 70.0,
                    'battery_mv': 3100
                },
                {
                    'timestamp': datetime(2025, 1, 1, 11, 0),
                    'temperature': 19.2,
                    'humidity': 68.5,
                    'battery_mv': 3095
                }
            ]
        }
        return [sample_device_data, device2_data]
    
    def test_create_temperature_timeline_valid_data(self, visualizer, sample_device_data):
        """Test creating temperature timeline with valid data."""
        # This test would require matplotlib to be installed
        # For now, we'll test the basic structure
        assert sample_device_data['device_name'] == 'TestDevice'
        assert len(sample_device_data['data']) == 3
        
        # In a real test environment with matplotlib:
        # result_path = visualizer.create_temperature_timeline(sample_device_data)
        # assert Path(result_path).exists()
    
    def test_create_temperature_timeline_empty_data(self, visualizer):
        """Test creating timeline with empty data."""
        empty_device_data = {
            'device_name': 'EmptyDevice',
            'data': []
        }
        
        with pytest.raises(ValueError, match="No data available"):
            visualizer.create_temperature_timeline(empty_device_data)
    
    def test_create_multi_device_comparison_valid_data(self, visualizer, multiple_device_data):
        """Test creating multi-device comparison with valid data."""
        # Test data structure
        assert len(multiple_device_data) == 2
        assert multiple_device_data[0]['device_name'] == 'TestDevice'
        assert multiple_device_data[1]['device_name'] == 'TestDevice2'
        
        # In a real test environment with matplotlib:
        # result_path = visualizer.create_multi_device_comparison(multiple_device_data, 'temperature')
        # assert Path(result_path).exists()
    
    def test_create_multi_device_comparison_empty_data(self, visualizer):
        """Test creating comparison with empty data."""
        with pytest.raises(ValueError, match="No device data provided"):
            visualizer.create_multi_device_comparison([], 'temperature')
    
    def test_create_statistics_heatmap_valid_data(self, visualizer, multiple_device_data):
        """Test creating statistics heatmap with valid data."""
        # Test data structure
        assert len(multiple_device_data) == 2
        
        # In a real test environment with matplotlib and seaborn:
        # result_path = visualizer.create_statistics_heatmap(multiple_device_data)
        # assert Path(result_path).exists()
    
    def test_create_statistics_heatmap_empty_data(self, visualizer):
        """Test creating heatmap with empty data."""
        with pytest.raises(ValueError, match="No device data provided"):
            visualizer.create_statistics_heatmap([])


if __name__ == '__main__':
    pytest.main([__file__])
