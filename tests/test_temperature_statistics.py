"""
Unit tests for temperature_statistics module.
"""

import pytest
import json
import tempfile
from pathlib import Path
from datetime import datetime, timedelta
from unittest.mock import patch

from src.temperature_statistics import TemperatureStatistics


class TestTemperatureStatistics:
    """Test cases for TemperatureStatistics class."""
    
    def create_test_database(self, temp_dir: Path, devices_data: dict) -> Path:
        """Create a temporary test database file."""
        db_path = temp_dir / "test_database.json"
        test_data = {
            "metadata": {
                "created": "2024-01-01T12:00:00",
                "last_updated": "2024-01-01T15:00:00",
                "total_records": sum(len(d["records"]) for d in devices_data.values())
            },
            "devices": devices_data
        }
        
        with open(db_path, 'w', encoding='utf-8') as f:
            json.dump(test_data, f, indent=2)
        
        return db_path
    
    def test_stat001_continuous_5min_intervals(self):
        """Test STAT001 with continuous 5-minute intervals."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Create test data with 5-minute intervals
            base_time = datetime(2024, 1, 1, 12, 0, 0)
            records = []
            for i in range(12):  # 1 hour of data
                timestamp = base_time + timedelta(minutes=i * 5)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 20.0 + i * 0.1,
                    "humidity": 50.0,
                    "battery": 3000
                })
            
            devices_data = {
                "TestDevice": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            # Test with default 35-minute max gap
            intervals = stats.stat001_active_time_intervals()
            
            assert "TestDevice" in intervals
            assert len(intervals["TestDevice"]) == 1  # Should be one continuous interval
            
            interval = intervals["TestDevice"][0]
            assert interval["start_time"] == base_time.strftime('%Y-%m-%d %H:%M:%S')
            assert interval["end_time"] == (base_time + timedelta(minutes=55)).strftime('%Y-%m-%d %H:%M:%S')  # Last record
            assert interval["record_count"] == 12
    
    def test_stat001_continuous_30min_intervals(self):
        """Test STAT001 with continuous 30-minute intervals (within 35-minute tolerance)."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Create test data with 30-minute intervals
            base_time = datetime(2024, 1, 1, 12, 0, 0)
            records = []
            for i in range(6):  # 2.5 hours of data
                timestamp = base_time + timedelta(minutes=i * 30)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 20.0 + i * 0.5,
                    "humidity": 50.0,
                    "battery": 3000
                })
            
            devices_data = {
                "TestDevice30": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            # Test with default 35-minute max gap
            intervals = stats.stat001_active_time_intervals()
            
            assert "TestDevice30" in intervals
            assert len(intervals["TestDevice30"]) == 1  # Should be one continuous interval
            
            interval = intervals["TestDevice30"][0]
            assert interval["start_time"] == base_time.strftime('%Y-%m-%d %H:%M:%S')
            assert interval["end_time"] == (base_time + timedelta(minutes=150)).strftime('%Y-%m-%d %H:%M:%S')  # Last record
            assert interval["record_count"] == 6
    
    def test_stat001_gap_exceeds_35min_threshold(self):
        """Test STAT001 with a gap exceeding 35-minute threshold."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            base_time = datetime(2024, 1, 1, 12, 0, 0)
            records = []
            
            # First interval: 3 records with 5-minute intervals
            for i in range(3):
                timestamp = base_time + timedelta(minutes=i * 5)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 20.0,
                    "humidity": 50.0,
                    "battery": 3000
                })
            
            # 40-minute gap (exceeds 35-minute threshold)
            gap_start = base_time + timedelta(minutes=10)
            gap_end = gap_start + timedelta(minutes=40)
            
            # Second interval: 3 records with 5-minute intervals
            for i in range(3):
                timestamp = gap_end + timedelta(minutes=i * 5)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 21.0,
                    "humidity": 51.0,
                    "battery": 2990
                })
            
            devices_data = {
                "TestDeviceGap": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            intervals = stats.stat001_active_time_intervals()
            
            assert "TestDeviceGap" in intervals
            assert len(intervals["TestDeviceGap"]) == 2  # Should be two separate intervals
            
            # First interval
            interval1 = intervals["TestDeviceGap"][0]
            assert interval1["start_time"] == base_time.strftime('%Y-%m-%d %H:%M:%S')
            assert interval1["end_time"] == (base_time + timedelta(minutes=10)).strftime('%Y-%m-%d %H:%M:%S')
            assert interval1["record_count"] == 3
            
            # Second interval
            interval2 = intervals["TestDeviceGap"][1]
            assert interval2["start_time"] == gap_end.strftime('%Y-%m-%d %H:%M:%S')
            assert interval2["end_time"] == (gap_end + timedelta(minutes=10)).strftime('%Y-%m-%d %H:%M:%S')
            assert interval2["record_count"] == 3
    
    def test_stat001_gap_within_35min_threshold(self):
        """Test STAT001 with a gap within 35-minute threshold (should remain continuous)."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            base_time = datetime(2024, 1, 1, 12, 0, 0)
            records = []
            
            # First part: 2 records
            for i in range(2):
                timestamp = base_time + timedelta(minutes=i * 5)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 20.0,
                    "humidity": 50.0,
                    "battery": 3000
                })
            
            # 30-minute gap (within 35-minute threshold)
            gap_start = base_time + timedelta(minutes=5)
            gap_end = gap_start + timedelta(minutes=30)
            
            # Second part: 2 records
            for i in range(2):
                timestamp = gap_end + timedelta(minutes=i * 5)
                records.append({
                    "timestamp": timestamp.isoformat(),
                    "temperature": 21.0,
                    "humidity": 51.0,
                    "battery": 2990
                })
            
            devices_data = {
                "TestDeviceSmallGap": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            intervals = stats.stat001_active_time_intervals()
            
            assert "TestDeviceSmallGap" in intervals
            assert len(intervals["TestDeviceSmallGap"]) == 1  # Should remain one continuous interval
            
            interval = intervals["TestDeviceSmallGap"][0]
            assert interval["start_time"] == base_time.strftime('%Y-%m-%d %H:%M:%S')
            assert interval["end_time"] == (gap_end + timedelta(minutes=5)).strftime('%Y-%m-%d %H:%M:%S')  # Last record
            assert interval["record_count"] == 4
    
    def test_stat001_empty_device(self):
        """Test STAT001 with empty device records."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            devices_data = {
                "EmptyDevice": {
                    "records": [],
                    "total_records": 0
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            intervals = stats.stat001_active_time_intervals()
            
            assert "EmptyDevice" in intervals
            assert len(intervals["EmptyDevice"]) == 0
    
    def test_stat001_invalid_timestamps(self):
        """Test STAT001 handles invalid timestamps gracefully."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            records = [
                {
                    "timestamp": "invalid-timestamp",
                    "temperature": 20.0,
                    "humidity": 50.0,
                    "battery": 3000
                },
                {
                    "timestamp": "2024-01-01T12:00:00",
                    "temperature": 20.5,
                    "humidity": 51.0,
                    "battery": 3000
                }
            ]
            
            devices_data = {
                "TestDeviceInvalid": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            # Should handle invalid timestamp and process valid one
            intervals = stats.stat001_active_time_intervals()
            
            assert "TestDeviceInvalid" in intervals
            assert len(intervals["TestDeviceInvalid"]) == 1
            assert intervals["TestDeviceInvalid"][0]["record_count"] == 1  # Only valid record
    
    def test_database_info(self):
        """Test get_database_info method."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            devices_data = {
                "Device1": {
                    "records": [{"timestamp": "2024-01-01T12:00:00", "temperature": 20.0, "humidity": 50.0, "battery": 3000}],
                    "total_records": 1
                },
                "Device2": {
                    "records": [{"timestamp": "2024-01-01T13:00:00", "temperature": 21.0, "humidity": 52.0, "battery": 2990}],
                    "total_records": 1
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            info = stats.get_database_info()
            
            assert info["total_devices"] == 2
            assert info["total_records"] == 2
            assert info["last_updated"] == "2024-01-01T15:00:00"
    
    def test_text_report_generation(self):
        """Test STAT001 text report generation."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            base_time = datetime(2024, 1, 1, 12, 0, 0)
            records = [
                {
                    "timestamp": base_time.isoformat(),
                    "temperature": 20.0,
                    "humidity": 50.0,
                    "battery": 3000
                },
                {
                    "timestamp": (base_time + timedelta(minutes=5)).isoformat(),
                    "temperature": 20.1,
                    "humidity": 50.1,
                    "battery": 3000
                }
            ]
            
            devices_data = {
                "TestDevice": {
                    "records": records,
                    "total_records": len(records)
                }
            }
            
            db_path = self.create_test_database(temp_path, devices_data)
            stats = TemperatureStatistics(str(db_path))
            
            intervals = stats.stat001_active_time_intervals()
            report = stats.stat001_format_text_report(intervals)
            
            assert "TestDevice" in report
            assert "2024-01-01 12:00:00" in report
            assert "2024-01-01 12:05:00" in report
            assert "2 / 2 expected" in report


if __name__ == "__main__":
    pytest.main([__file__])
