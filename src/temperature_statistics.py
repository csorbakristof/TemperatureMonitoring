"""
Statistics Module - Temperature Data Analysis

This module provides various statistics for temperature monitoring data,
starting with STAT001 - Active time interval of devices.
"""

import json
import logging
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional

logger = logging.getLogger(__name__)


class TemperatureStatistics:
    """Provides statistical analysis of temperature monitoring data."""
    
    def __init__(self, json_db_path: str = "data/temperature_database.json"):
        self.json_db_path = Path(json_db_path)
        self.database = self._load_database()
        
    def _load_database(self) -> Dict:
        """Load the JSON database."""
        if not self.json_db_path.exists():
            raise FileNotFoundError(f"Database not found: {self.json_db_path}")
        
        try:
            with open(self.json_db_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON database: {e}")
    
    def stat001_active_time_intervals(self, expected_interval_minutes: int = 5, 
                                    max_gap_minutes: int = 35) -> Dict[str, List[Dict]]:
        """
        STAT001: Active time interval of devices
        
        Shows every available device name and the time intervals we have data for.
        If the time interval is not continuous (some devices record every 5 minutes, others every 30 minutes),
        multiple intervals are listed, but only if the missing data gap is more than 35 minutes.
        
        Args:
            expected_interval_minutes: Expected interval between records (default 5 minutes)
            max_gap_minutes: Maximum gap before splitting intervals (default 35 minutes)
            
        Returns:
            Dictionary with device names as keys and list of time intervals as values
        """
        if not self.database.get('devices'):
            return {}
        
        results = {}
        max_gap = timedelta(minutes=max_gap_minutes)
        expected_interval = timedelta(minutes=expected_interval_minutes)
        
        for device_name, device_data in self.database['devices'].items():
            records = device_data.get('records', [])
            
            if not records:
                results[device_name] = []
                continue
            
            # Convert timestamp strings to datetime objects and sort
            timestamps = []
            for record in records:
                try:
                    timestamp = datetime.fromisoformat(record['timestamp'])
                    timestamps.append(timestamp)
                except (ValueError, KeyError) as e:
                    logger.warning(f"Invalid timestamp in device {device_name}: {e}")
                    continue
            
            timestamps.sort()
            
            if not timestamps:
                results[device_name] = []
                continue
            
            # Find continuous intervals
            intervals = []
            current_start = timestamps[0]
            current_end = timestamps[0]
            
            for i in range(1, len(timestamps)):
                time_gap = timestamps[i] - timestamps[i-1]
                
                if time_gap <= max_gap:
                    # Continue current interval
                    current_end = timestamps[i]
                else:
                    # Gap too large - finish current interval and start new one
                    intervals.append({
                        'start': current_start,
                        'end': current_end,
                        'duration': current_end - current_start,
                        'record_count': self._count_records_in_interval(timestamps, current_start, current_end)
                    })
                    
                    current_start = timestamps[i]
                    current_end = timestamps[i]
            
            # Add the last interval
            intervals.append({
                'start': current_start,
                'end': current_end,
                'duration': current_end - current_start,
                'record_count': self._count_records_in_interval(timestamps, current_start, current_end)
            })
            
            # Format intervals for output
            formatted_intervals = []
            for interval in intervals:
                formatted_intervals.append({
                    'start_time': interval['start'].strftime('%Y-%m-%d %H:%M:%S'),
                    'end_time': interval['end'].strftime('%Y-%m-%d %H:%M:%S'),
                    'duration_hours': round(interval['duration'].total_seconds() / 3600, 2),
                    'duration_days': round(interval['duration'].total_seconds() / 86400, 1),
                    'record_count': interval['record_count'],
                    'expected_records': int(interval['duration'].total_seconds() / (expected_interval_minutes * 60)) + 1,
                    'completeness_percent': round(
                        (interval['record_count'] / max(1, int(interval['duration'].total_seconds() / (expected_interval_minutes * 60)) + 1)) * 100, 
                        1
                    )
                })
            
            results[device_name] = formatted_intervals
        
        return results
    
    def _count_records_in_interval(self, timestamps: List[datetime], start: datetime, end: datetime) -> int:
        """Count how many timestamps fall within the given interval."""
        return sum(1 for ts in timestamps if start <= ts <= end)
    
    def stat001_format_text_report(self, intervals_data: Dict[str, List[Dict]]) -> str:
        """
        Format STAT001 results as a text report.
        
        Args:
            intervals_data: Result from stat001_active_time_intervals()
            
        Returns:
            Formatted text report
        """
        if not intervals_data:
            return "No device data available."
        
        report_lines = []
        report_lines.append("STAT001: Active Time Intervals of Devices")
        report_lines.append("=" * 60)
        report_lines.append("")
        
        for device_name, intervals in intervals_data.items():
            report_lines.append(f"Device: {device_name}")
            report_lines.append("-" * 40)
            
            if not intervals:
                report_lines.append("  No data intervals found")
            else:
                for i, interval in enumerate(intervals, 1):
                    report_lines.append(f"  Interval {i}:")
                    report_lines.append(f"    Start:      {interval['start_time']}")
                    report_lines.append(f"    End:        {interval['end_time']}")
                    report_lines.append(f"    Duration:   {interval['duration_days']} days ({interval['duration_hours']} hours)")
                    report_lines.append(f"    Records:    {interval['record_count']} / {interval['expected_records']} expected ({interval['completeness_percent']}% complete)")
            
            report_lines.append("")
        
        # Add summary
        total_devices = len(intervals_data)
        total_intervals = sum(len(intervals) for intervals in intervals_data.values())
        total_records = sum(
            sum(interval['record_count'] for interval in intervals) 
            for intervals in intervals_data.values()
        )
        
        report_lines.append("Summary:")
        report_lines.append(f"  Total devices: {total_devices}")
        report_lines.append(f"  Total intervals: {total_intervals}")
        report_lines.append(f"  Total records: {total_records}")
        
        return "\n".join(report_lines)
    
    def stat002_temperature_gradient_ventilation_analysis(self) -> Dict[str, Any]:
        """
        STAT002: Simplified temperature difference analysis for ventilation status.
        
        This function focuses on the most useful indicator: temperature difference
        between T3_Kek (room) and T1_BE (intake). This provides the clearest
        physical indication of ventilation system behavior.
        
        Returns:
            Dictionary with simplified ventilation analysis results
        """
        if not self.database.get('devices'):
            return {"error": "No device data available"}
        
        # Required devices for ventilation analysis
        required_devices = ['T1_BE', 'T3_Kek', 'T2_Terasz']
        devices = self.database['devices']
        
        # Check if required devices exist
        missing_devices = [dev for dev in required_devices if dev not in devices]
        if missing_devices:
            return {
                "error": f"Missing required devices for ventilation analysis: {missing_devices}",
                "available_devices": list(devices.keys())
            }
        
        # Get synchronized data points
        sync_data = self._synchronize_device_data(required_devices)
        
        if len(sync_data) < 10:
            return {"error": f"Insufficient synchronized data points for analysis: {len(sync_data)} found (minimum 10 required)"}
        
        # Convert to pandas DataFrame for easier analysis
        import pandas as pd
        df = pd.DataFrame(sync_data)
        df['timestamp'] = pd.to_datetime(df['timestamp'], format='ISO8601')
        df = df.sort_values('timestamp').reset_index(drop=True)
        
        # Calculate temperature differences - this is the key insight
        results = []
        
        for i in range(len(df)):
            current_time = df.loc[i, 'timestamp']
            t1_current = df.loc[i, 'T1_BE']      # Intake
            t2_current = df.loc[i, 'T2_Terasz']  # External
            t3_current = df.loc[i, 'T3_Kek']     # Room
            
            # The most useful metric: Room - Intake temperature difference
            temp_difference = t3_current - t1_current
            
            results.append({
                'timestamp': current_time.isoformat(),
                'T1_BE': round(t1_current, 2),
                'T2_Terasz': round(t2_current, 2),
                'T3_Kek': round(t3_current, 2),
                'temperature_difference': round(temp_difference, 2)
            })
        
        # Calculate summary statistics for temperature difference
        temp_diffs = [r['temperature_difference'] for r in results]
        mean_diff = float(np.mean(temp_diffs))
        std_diff = float(np.std(temp_diffs))
        min_diff = float(np.min(temp_diffs))
        max_diff = float(np.max(temp_diffs))
        
        return {
            'analysis_method': 'Temperature Difference Analysis (Simplified)',
            'description': 'Room (T3_Kek) - Intake (T1_BE) temperature difference over time',
            'total_data_points': len(results),
            'time_range': {
                'start': results[0]['timestamp'] if results else None,
                'end': results[-1]['timestamp'] if results else None
            },
            'temperature_difference_statistics': {
                'mean_celsius': round(mean_diff, 2),
                'std_celsius': round(std_diff, 2),
                'min_celsius': round(min_diff, 2),
                'max_celsius': round(max_diff, 2),
                'range_celsius': round(max_diff - min_diff, 2)
            },
            'interpretation': {
                'positive_values': 'Room warmer than intake (likely ventilation OFF)',
                'negative_values': 'Intake warmer than room (possible ventilation ON or external influence)',
                'near_zero_values': 'Similar temperatures between room and intake'
            },
            'detailed_results': results
        }
    
    def _synchronize_device_data(self, device_names: List[str], tolerance_minutes: int = 15) -> List[Dict[str, Any]]:
        """
        Synchronize data from multiple devices by timestamp with tolerance.
        
        Args:
            device_names: List of device names to synchronize
            tolerance_minutes: Maximum time difference allowed for synchronization (default: 15 minutes)
            
        Returns:
            List of synchronized data points with all device readings
        """
        devices = self.database['devices']
        tolerance = timedelta(minutes=tolerance_minutes)
        
        # Collect all records from all devices with timestamps
        all_records = []
        
        for device_name in device_names:
            if device_name not in devices:
                continue
                
            records = devices[device_name].get('records', [])
            for record in records:
                try:
                    timestamp_str = record['timestamp']
                    timestamp = datetime.fromisoformat(timestamp_str)
                    temp = float(record['temperature'])
                    
                    all_records.append({
                        'device': device_name,
                        'timestamp': timestamp,
                        'temperature': temp
                    })
                    
                except (ValueError, KeyError) as e:
                    logger.warning(f"Invalid record in device {device_name}: {e}")
                    continue
        
        # Sort all records by timestamp
        all_records.sort(key=lambda x: x['timestamp'])
        
        if not all_records:
            return []
        
        # Group records by approximate time windows
        synchronized_data = []
        current_window = []
        window_start_time = all_records[0]['timestamp']
        
        for record in all_records:
            # Check if this record is within the current time window
            if record['timestamp'] - window_start_time <= tolerance:
                current_window.append(record)
            else:
                # Process current window if it has data from all required devices
                window_data = self._process_time_window(current_window, device_names)
                if window_data:
                    synchronized_data.append(window_data)
                
                # Start new window
                current_window = [record]
                window_start_time = record['timestamp']
        
        # Process the last window
        window_data = self._process_time_window(current_window, device_names)
        if window_data:
            synchronized_data.append(window_data)
        
        return synchronized_data
    
    def _process_time_window(self, window_records: List[Dict], required_devices: List[str]) -> Optional[Dict[str, Any]]:
        """
        Process a time window to create a synchronized data point.
        
        Args:
            window_records: All records within the time window
            required_devices: List of required device names
            
        Returns:
            Synchronized data point or None if not all devices are present
        """
        # Group by device
        device_temps = {}
        timestamps = []
        
        for record in window_records:
            device = record['device']
            temp = record['temperature']
            timestamp = record['timestamp']
            
            # Take the most recent temperature for each device in this window
            if device not in device_temps or timestamp > device_temps[device]['timestamp']:
                device_temps[device] = {
                    'temperature': temp,
                    'timestamp': timestamp
                }
            
            timestamps.append(timestamp)
        
        # Check if all required devices are present
        if not all(device in device_temps for device in required_devices):
            return None
        
        # Create synchronized data point using average timestamp
        avg_timestamp = min(timestamps) + (max(timestamps) - min(timestamps)) / 2
        
        data_point = {'timestamp': avg_timestamp.isoformat()}
        for device in required_devices:
            data_point[device] = device_temps[device]['temperature']
        
        return data_point

    def get_database_info(self) -> Dict[str, Any]:
        """Get basic information about the database."""
        return {
            "total_devices": len(self.database.get('devices', {})),
            "total_records": self.database.get('metadata', {}).get('total_records', 0),
            "last_updated": self.database.get('metadata', {}).get('last_updated', 'Unknown'),
            "device_names": list(self.database.get('devices', {}).keys())
        }


def main():
    """Main function for statistics analysis."""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        stats = TemperatureStatistics()
        
        print("Temperature Data Statistics")
        print("=" * 50)
        
        # Show database info
        db_info = stats.get_database_info()
        print(f"Database contains {db_info['total_devices']} devices with {db_info['total_records']} total records")
        print(f"Last updated: {db_info['last_updated']}")
        print("")
        
        # Run STAT001
        print("Running STAT001: Active Time Intervals...")
        intervals_data = stats.stat001_active_time_intervals()
        
        # Generate and print text report
        report = stats.stat001_format_text_report(intervals_data)
        print(report)
        
        # Save report to file
        report_path = Path("output/STAT001_active_intervals_report.txt")
        report_path.parent.mkdir(exist_ok=True)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        print(f"\nReport saved to: {report_path}")
        
        # Run STAT002
        print("\n" + "=" * 50)
        print("Running STAT002: Simplified Temperature Difference Analysis...")
        ventilation_data = stats.stat002_temperature_gradient_ventilation_analysis()
        
        if 'error' in ventilation_data:
            print(f"STAT002 Error: {ventilation_data['error']}")
            if 'available_devices' in ventilation_data:
                print(f"Available devices: {ventilation_data['available_devices']}")
        else:
            # Print summary
            temp_stats = ventilation_data['temperature_difference_statistics']
            print(f"Analysis method: {ventilation_data['analysis_method']}")
            print(f"Description: {ventilation_data['description']}")
            print(f"Total data points analyzed: {ventilation_data['total_data_points']}")
            print(f"Time range: {ventilation_data['time_range']['start']} to {ventilation_data['time_range']['end']}")
            print("")
            print("Temperature Difference Statistics (Room - Intake):")
            print(f"  Mean:  {temp_stats['mean_celsius']}°C")
            print(f"  Std:   {temp_stats['std_celsius']}°C")
            print(f"  Range: {temp_stats['min_celsius']}°C to {temp_stats['max_celsius']}°C")
            print(f"  Total Range: {temp_stats['range_celsius']}°C")
            print("")
            print("Interpretation:")
            for key, value in ventilation_data['interpretation'].items():
                print(f"  {key.replace('_', ' ').title()}: {value}")
            
            # Save detailed results to JSON
            ventilation_report_path = Path("output/STAT002_ventilation_analysis.json")
            with open(ventilation_report_path, 'w', encoding='utf-8') as f:
                json.dump(ventilation_data, f, indent=2)
            
            print(f"\nDetailed results saved to: {ventilation_report_path}")
            print("\nRun 'python visualize_ventilation.py' to create the temperature difference heatmap.")
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("Please run the data import first using data_importer.py")
    except Exception as e:
        print(f"Error: {e}")
        logger.error(f"Statistics error: {e}")


if __name__ == "__main__":
    main()
