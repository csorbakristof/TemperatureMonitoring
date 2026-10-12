#!/usr/bin/env python3
"""
Test for STAT002 Temperature Gradient Ventilation Analysis
"""

import sys
import unittest
from pathlib import Path

# Add src directory to path
sys.path.append(str(Path(__file__).parent / "src"))

from temperature_statistics import TemperatureStatistics

class TestSTAT002VentilationAnalysis(unittest.TestCase):
    
    def setUp(self):
        """Set up test environment."""
        self.stats = TemperatureStatistics()
        
    def test_ventilation_analysis_runs(self):
        """Test that STAT002 analysis runs without errors."""
        try:
            result = self.stats.stat002_temperature_gradient_ventilation_analysis()
            
            # Should not return an error
            self.assertNotIn('error', result, "STAT002 should run without errors")
            
            # Should contain expected keys
            expected_keys = ['analysis_method', 'total_data_points', 'summary', 'detailed_results']
            for key in expected_keys:
                self.assertIn(key, result, f"Result should contain {key}")
                
            # Summary should have the right structure
            summary = result['summary']
            summary_keys = ['ventilation_on_periods', 'ventilation_off_periods', 'uncertain_periods',
                          'on_percentage', 'off_percentage', 'uncertain_percentage', 'average_confidence']
            for key in summary_keys:
                self.assertIn(key, summary, f"Summary should contain {key}")
                
        except Exception as e:
            self.fail(f"STAT002 analysis failed with exception: {e}")
    
    def test_synchronization_logic(self):
        """Test that device synchronization works."""
        required_devices = ['T1_BE', 'T3_Kek', 'T2_Terasz']
        
        # Test synchronization
        sync_data = self.stats._synchronize_device_data(required_devices)
        
        # Should return some data points
        self.assertGreater(len(sync_data), 0, "Should find synchronized data points")
        
        # Each data point should have all required devices
        if sync_data:
            first_point = sync_data[0]
            for device in required_devices:
                self.assertIn(device, first_point, f"Data point should contain {device}")
                
            # Should have timestamp
            self.assertIn('timestamp', first_point, "Data point should have timestamp")
    
    def test_ventilation_parameters(self):
        """Test ventilation analysis with different parameters."""
        # Test with different window sizes
        result_60min = self.stats.stat002_temperature_gradient_ventilation_analysis(window_minutes=60)
        result_30min = self.stats.stat002_temperature_gradient_ventilation_analysis(window_minutes=30)
        
        if 'error' not in result_60min and 'error' not in result_30min:
            # Different window sizes should potentially give different results
            self.assertIsInstance(result_60min['total_data_points'], int)
            self.assertIsInstance(result_30min['total_data_points'], int)
            
        # Test with different gradient thresholds  
        result_low_threshold = self.stats.stat002_temperature_gradient_ventilation_analysis(gradient_threshold=0.2)
        result_high_threshold = self.stats.stat002_temperature_gradient_ventilation_analysis(gradient_threshold=1.0)
        
        if 'error' not in result_low_threshold and 'error' not in result_high_threshold:
            # Should run with different thresholds
            self.assertIsInstance(result_low_threshold['total_data_points'], int)
            self.assertIsInstance(result_high_threshold['total_data_points'], int)
    
    def test_missing_devices(self):
        """Test behavior when required devices are missing."""
        # Test with non-existent devices
        self.stats.database = {'devices': {'T1_BE': {'records': []}}}  # Only partial data
        
        result = self.stats.stat002_temperature_gradient_ventilation_analysis()
        
        # Should return error about missing devices
        self.assertIn('error', result, "Should return error for missing devices")
        self.assertIn('Missing required devices', result['error'])

if __name__ == '__main__':
    print("Running STAT002 Ventilation Analysis Tests...")
    unittest.main()
