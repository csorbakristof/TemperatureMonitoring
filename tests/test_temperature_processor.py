"""
Unit tests for the temperature processor module.
"""

import pytest
import tempfile
import zipfile
import csv
from datetime import datetime
from pathlib import Path
import sys

# Add src directory to path
sys.path.append(str(Path(__file__).parent.parent / 'src'))

from temperature_processor import TemperatureDataProcessor


class TestTemperatureDataProcessor:
    """Test cases for TemperatureDataProcessor class."""
    
    @pytest.fixture
    def processor(self):
        """Create a processor instance for testing."""
        return TemperatureDataProcessor()
    
    @pytest.fixture
    def sample_csv_content(self):
        """Sample CSV content for testing."""
        return [
            "T3_TestDevice",
            "Time-Data=(A2/86400)+25569;Temp;Humi;Vbat",
            "1609459200;13.72;86.58;2943",  # 2021-01-01 00:00:00 UTC
            "1609459500;13.57;86.7;2957",   # 2021-01-01 00:05:00 UTC
            "1609459800;13.5;86.88;2960"    # 2021-01-01 00:10:00 UTC
        ]
    
    @pytest.fixture
    def sample_csv_file(self, sample_csv_content):
        """Create a temporary CSV file for testing."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            f.write('\n'.join(sample_csv_content))
            return f.name
    
    @pytest.fixture
    def sample_zip_file(self, sample_csv_content):
        """Create a temporary ZIP file containing CSV data."""
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as zip_file:
            with zipfile.ZipFile(zip_file.name, 'w') as zf:
                zf.writestr('test_device.csv', '\n'.join(sample_csv_content))
            return zip_file.name
    
    def test_csv_timestamp_to_datetime_conversion(self, processor):
        """Test CSV timestamp to datetime conversion."""
        # Test known Unix timestamp: 1609459200 = 2021-01-01 00:00:00 UTC
        csv_timestamp = 1609459200.0
        result = processor._csv_timestamp_to_datetime(csv_timestamp)
        
        assert isinstance(result, datetime)
        assert result.year == 2021
        assert result.month == 1
        assert result.day == 1
    
    def test_parse_csv_file_valid(self, processor, sample_csv_file):
        """Test parsing a valid CSV file."""
        result = processor.parse_csv_file(sample_csv_file)
        
        assert result['device_name'] == 'T3_TestDevice'
        assert len(result['data']) == 3
        assert result['file_path'] == sample_csv_file
        
        # Check first record
        first_record = result['data'][0]
        assert isinstance(first_record['timestamp'], datetime)
        assert first_record['temperature'] == 13.72
        assert first_record['humidity'] == 86.58
        assert first_record['battery_mv'] == 2943
    
    def test_parse_csv_file_invalid_format(self, processor):
        """Test parsing CSV file with invalid format."""
        # Create CSV with insufficient lines
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            f.write("Device\nHeader")  # Only 2 lines
            
        with pytest.raises(ValueError, match="insufficient data"):
            processor.parse_csv_file(f.name)
    
    def test_parse_csv_file_invalid_data(self, processor):
        """Test parsing CSV file with invalid data lines."""
        content = [
            "T3_TestDevice",
            "Time-Data=(A2/86400)+25569;Temp;Humi;Vbat",
            "invalid;data;line",  # Invalid line
            "1791;13.72;86.58;2943"  # Valid line
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            f.write('\n'.join(content))
            
        result = processor.parse_csv_file(f.name)
        
        # Should skip invalid line and process valid one
        assert len(result['data']) == 1
        assert result['data'][0]['temperature'] == 13.72
    
    def test_extract_zip_files(self, processor, sample_zip_file):
        """Test extracting CSV files from ZIP archive."""
        with tempfile.TemporaryDirectory() as temp_dir:
            extracted_files = processor.extract_zip_files(sample_zip_file, temp_dir)
            
            assert len(extracted_files) == 1
            assert extracted_files[0].name == 'test_device.csv'
            assert extracted_files[0].exists()
    
    def test_extract_zip_files_invalid_zip(self, processor):
        """Test extracting from invalid ZIP file."""
        # Create a non-ZIP file
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as f:
            f.write(b"Not a zip file")
            
        with pytest.raises(zipfile.BadZipFile):
            processor.extract_zip_files(f.name)
    
    def test_process_zip_file(self, processor, sample_zip_file):
        """Test complete ZIP file processing."""
        result = processor.process_zip_file(sample_zip_file)
        
        assert len(result) == 1
        device_data = result[0]
        
        assert device_data['device_name'] == 'T3_TestDevice'
        assert len(device_data['data']) == 3
        
        # Verify data structure
        for record in device_data['data']:
            assert 'timestamp' in record
            assert 'temperature' in record
            assert 'humidity' in record
            assert 'battery_mv' in record
    
    def test_process_zip_file_nonexistent(self, processor):
        """Test processing non-existent ZIP file."""
        with pytest.raises(FileNotFoundError):
            processor.process_zip_file("nonexistent.zip")
    
    def test_process_zip_file_no_csv(self, processor):
        """Test processing ZIP file with no CSV files."""
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as zip_file:
            with zipfile.ZipFile(zip_file.name, 'w') as zf:
                zf.writestr('readme.txt', 'No CSV files here')
            
            result = processor.process_zip_file(zip_file.name)
            assert len(result) == 0
    
    def test_multiple_csv_files_in_zip(self, processor):
        """Test processing ZIP file with multiple CSV files."""
        csv_content1 = [
            "Device_1",
            "Time-Data=(A2/86400)+25569;Temp;Humi;Vbat",
            "1791;13.72;86.58;2943"
        ]
        
        csv_content2 = [
            "Device_2",
            "Time-Data=(A2/86400)+25569;Temp;Humi;Vbat",
            "1791;15.50;80.25;2800"
        ]
        
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as zip_file:
            with zipfile.ZipFile(zip_file.name, 'w') as zf:
                zf.writestr('device1.csv', '\n'.join(csv_content1))
                zf.writestr('device2.csv', '\n'.join(csv_content2))
            
            result = processor.process_zip_file(zip_file.name)
            
            assert len(result) == 2
            device_names = [d['device_name'] for d in result]
            assert 'Device_1' in device_names
            assert 'Device_2' in device_names


if __name__ == '__main__':
    pytest.main([__file__])
