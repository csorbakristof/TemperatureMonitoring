"""
Unit tests for the Excel exporter module.
"""

import pytest
import tempfile
import pandas as pd
from datetime import datetime
from pathlib import Path
import sys

# Add src directory to path
sys.path.append(str(Path(__file__).parent.parent / 'src'))

from excel_exporter import ExcelExporter


class TestExcelExporter:
    """Test cases for ExcelExporter class."""
    
    @pytest.fixture
    def exporter(self):
        """Create an exporter instance for testing."""
        with tempfile.TemporaryDirectory() as temp_dir:
            return ExcelExporter(output_dir=temp_dir)
    
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
    
    def test_calculate_statistics(self, exporter, sample_device_data):
        """Test statistics calculation."""
        df = pd.DataFrame(sample_device_data['data'])
        stats = exporter._calculate_statistics(df)
        
        assert 'Temperature (°C)' in stats.columns
        assert 'Humidity (%)' in stats.columns
        assert 'Battery (mV)' in stats.columns
        
        # Check temperature statistics
        temp_stats = stats['Temperature (°C)']
        assert temp_stats['Mean'] == pytest.approx(21.27, rel=1e-2)
        assert temp_stats['Min'] == 20.5
        assert temp_stats['Max'] == 22.1
    
    def test_create_hourly_averages(self, exporter, sample_device_data):
        """Test hourly averages creation."""
        df = pd.DataFrame(sample_device_data['data'])
        hourly = exporter._create_hourly_averages(df)
        
        assert len(hourly) == 3  # 3 hourly periods
        assert 'timestamp' in hourly.columns
        assert 'temperature' in hourly.columns
        assert 'humidity' in hourly.columns
        assert 'battery_mv' in hourly.columns
    
    def test_export_device_data_empty_data(self, exporter):
        """Test exporting device with empty data."""
        empty_device_data = {
            'device_name': 'EmptyDevice',
            'data': []
        }
        
        with pytest.raises(ValueError, match="No data available"):
            exporter.export_device_data(empty_device_data)
    
    def test_export_all_devices_empty_data(self, exporter):
        """Test exporting with empty device list."""
        with pytest.raises(ValueError, match="No device data provided"):
            exporter.export_all_devices([])
    
    # Note: Full Excel export tests would require openpyxl and pandas to be installed
    # In a real test environment, you would test:
    # - Actual file creation
    # - Excel file structure
    # - Data integrity in Excel sheets
    # - Formatting verification


if __name__ == '__main__':
    pytest.main([__file__])
