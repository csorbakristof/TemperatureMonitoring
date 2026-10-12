#!/usr/bin/env python3
"""
Calendar Image Generator

Creates calendar heatmap visualizations for temperature data and heating activity.
Each row represents a day, each column represents a 5-minute interval within the day.
"""

import json
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.patches import Rectangle
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd
from typing import Dict, List, Tuple, Optional
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
SAMPLE_INTERVAL_MINUTES = 5
SAMPLES_PER_DAY = 24 * 60 // SAMPLE_INTERVAL_MINUTES  # 288 samples per day
ZONE_DEVICES = {
    'T8_Z1': 'Zone 1',
    'T6_Z2': 'Zone 2'
}
TEMP_DIFF_DEVICES = ['T3_Kek', 'T2_Terasz']
OUTSIDE_DEVICE = 'T2_Terasz'


class CalendarImageGenerator:
    """Generates calendar heatmap images for temperature and heating data."""
    
    def __init__(self, 
                 temperature_db_path: str = "data/temperature_database.json",
                 heating_cycles_path: str = "output/heating_cycles.json",
                 output_dir: str = "output"):
        self.temperature_db_path = Path(temperature_db_path)
        self.heating_cycles_path = Path(heating_cycles_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        self.temperature_db = self._load_json(self.temperature_db_path)
        self.heating_cycles = self._load_json(self.heating_cycles_path)
        
    def _load_json(self, path: Path) -> Dict:
        """Load JSON file."""
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _get_date_range(self, device_name: str) -> Tuple[datetime, datetime]:
        """Get the full date range for a device."""
        records = self.temperature_db.get('devices', {}).get(device_name, {}).get('records', [])
        
        if not records:
            raise ValueError(f"No records found for device {device_name}")
        
        timestamps = [pd.to_datetime(r['timestamp']) for r in records]
        min_date = min(timestamps).date()
        max_date = max(timestamps).date()
        
        return min_date, max_date
    
    def _create_empty_calendar_matrix(self, start_date: datetime.date, end_date: datetime.date) -> np.ndarray:
        """Create an empty calendar matrix filled with NaN."""
        num_days = (end_date - start_date).days + 1
        matrix = np.full((num_days, SAMPLES_PER_DAY), np.nan)
        return matrix
    
    def _get_day_and_sample_index(self, timestamp: datetime, start_date: datetime.date) -> Tuple[int, int]:
        """Convert timestamp to day index and sample index within the day."""
        ts_date = timestamp.date()
        day_index = (ts_date - start_date).days
        
        # Calculate sample index within the day (0-287 for 5-minute intervals)
        minutes_since_midnight = timestamp.hour * 60 + timestamp.minute
        sample_index = minutes_since_midnight // SAMPLE_INTERVAL_MINUTES
        
        return day_index, sample_index
    
    def generate_temperature_calendar(self, device_name: str) -> str:
        """Generate calendar heatmap for temperature data."""
        logger.info(f"Generating temperature calendar for {device_name}...")
        
        # Get date range and create matrix
        start_date, end_date = self._get_date_range(device_name)
        matrix = self._create_empty_calendar_matrix(start_date, end_date)
        
        # Fill matrix with temperature data
        records = self.temperature_db['devices'][device_name]['records']
        for record in records:
            timestamp = pd.to_datetime(record['timestamp'])
            temperature = record['temperature']
            
            day_idx, sample_idx = self._get_day_and_sample_index(timestamp, start_date)
            if 0 <= day_idx < matrix.shape[0] and 0 <= sample_idx < matrix.shape[1]:
                matrix[day_idx, sample_idx] = temperature
        
        # Create the heatmap
        zone_id = device_name.split('_')[-1]  # Extract Z1 or Z2
        output_path = self.output_dir / f"TempCal_{zone_id}.png"
        
        self._create_heatmap(
            matrix=matrix,
            start_date=start_date,
            end_date=end_date,
            title=f"Temperature Calendar - {device_name} ({ZONE_DEVICES.get(device_name, 'Unknown')})",
            colormap='viridis',
            missing_color='white',
            output_path=output_path,
            value_label='Temperature (°C)'
        )
        
        logger.info(f"Saved temperature calendar: {output_path}")
        return str(output_path)
    
    def generate_heating_calendar(self, device_name: str) -> str:
        """Generate calendar heatmap for heating activity (binary)."""
        logger.info(f"Generating heating calendar for {device_name}...")
        
        # Get date range and create matrix (filled with 0 = no heating)
        start_date, end_date = self._get_date_range(device_name)
        matrix = self._create_empty_calendar_matrix(start_date, end_date)
        matrix[:] = 0  # Initialize with 0 (no heating)
        
        # Mark heating periods as 1
        cycles = self.heating_cycles.get(device_name, [])
        for cycle in cycles:
            start_time = pd.to_datetime(cycle['start'])
            end_time = pd.to_datetime(cycle['end'])
            
            # Mark all 5-minute intervals in this heating cycle
            current_time = start_time
            while current_time <= end_time:
                day_idx, sample_idx = self._get_day_and_sample_index(current_time, start_date)
                if 0 <= day_idx < matrix.shape[0] and 0 <= sample_idx < matrix.shape[1]:
                    matrix[day_idx, sample_idx] = 1
                
                current_time += timedelta(minutes=SAMPLE_INTERVAL_MINUTES)
        
        # Create the heatmap
        zone_id = device_name.split('_')[-1]  # Extract Z1 or Z2
        output_path = self.output_dir / f"Heating_{zone_id}.png"
        
        self._create_heatmap(
            matrix=matrix,
            start_date=start_date,
            end_date=end_date,
            title=f"Heating Activity Calendar - {device_name} ({ZONE_DEVICES.get(device_name, 'Unknown')})",
            colormap='heating',  # Custom red/white colormap
            missing_color='gray',
            output_path=output_path,
            value_label='Heating Active',
            is_binary=True
        )
        
        logger.info(f"Saved heating calendar: {output_path}")
        return str(output_path)
    
    def generate_temperature_difference_calendar(self) -> str:
        """Generate calendar heatmap for temperature difference between T3_Kek and T2_Terasz."""
        logger.info(f"Generating temperature difference calendar for {TEMP_DIFF_DEVICES[0]} - {TEMP_DIFF_DEVICES[1]}...")
        
        device1_name = TEMP_DIFF_DEVICES[0]  # T3_Kek
        device2_name = TEMP_DIFF_DEVICES[1]  # T2_Terasz
        
        # Get date range covering both devices
        start_date1, end_date1 = self._get_date_range(device1_name)
        start_date2, end_date2 = self._get_date_range(device2_name)
        
        start_date = min(start_date1, start_date2)
        end_date = max(end_date1, end_date2)
        
        # Create matrices for both devices
        matrix1 = self._create_empty_calendar_matrix(start_date, end_date)
        matrix2 = self._create_empty_calendar_matrix(start_date, end_date)
        
        # Fill matrix1 with T3_Kek data
        records1 = self.temperature_db['devices'][device1_name]['records']
        for record in records1:
            timestamp = pd.to_datetime(record['timestamp'])
            temperature = record['temperature']
            
            day_idx, sample_idx = self._get_day_and_sample_index(timestamp, start_date)
            if 0 <= day_idx < matrix1.shape[0] and 0 <= sample_idx < matrix1.shape[1]:
                matrix1[day_idx, sample_idx] = temperature
        
        # Fill matrix2 with T2_Terasz data
        records2 = self.temperature_db['devices'][device2_name]['records']
        for record in records2:
            timestamp = pd.to_datetime(record['timestamp'])
            temperature = record['temperature']
            
            day_idx, sample_idx = self._get_day_and_sample_index(timestamp, start_date)
            if 0 <= day_idx < matrix2.shape[0] and 0 <= sample_idx < matrix2.shape[1]:
                matrix2[day_idx, sample_idx] = temperature
        
        # Calculate difference: T3_Kek - T2_Terasz
        # Only calculate where both values are present
        diff_matrix = np.full_like(matrix1, np.nan)
        valid_mask = ~np.isnan(matrix1) & ~np.isnan(matrix2)
        diff_matrix[valid_mask] = matrix1[valid_mask] - matrix2[valid_mask]
        
        # Create the heatmap
        output_path = self.output_dir / "TempDiff.png"
        
        self._create_heatmap(
            matrix=diff_matrix,
            start_date=start_date,
            end_date=end_date,
            title=f"Temperature Difference Calendar - {device1_name} - {device2_name}",
            colormap='viridis',
            missing_color='white',
            output_path=output_path,
            value_label='Temperature Difference (°C)'
        )
        
        logger.info(f"Saved temperature difference calendar: {output_path}")
        return str(output_path)
    
    def generate_outside_temperature_calendar(self) -> str:
        """Generate calendar heatmap for outside temperature (T2_Terasz)."""
        logger.info(f"Generating outside temperature calendar for {OUTSIDE_DEVICE}...")
        
        device_name = OUTSIDE_DEVICE
        
        # Get date range and create matrix
        start_date, end_date = self._get_date_range(device_name)
        matrix = self._create_empty_calendar_matrix(start_date, end_date)
        
        # Fill matrix with temperature data
        records = self.temperature_db['devices'][device_name]['records']
        for record in records:
            timestamp = pd.to_datetime(record['timestamp'])
            temperature = record['temperature']
            
            day_idx, sample_idx = self._get_day_and_sample_index(timestamp, start_date)
            if 0 <= day_idx < matrix.shape[0] and 0 <= sample_idx < matrix.shape[1]:
                matrix[day_idx, sample_idx] = temperature
        
        # Create the heatmap
        output_path = self.output_dir / "Temp_Outside.png"
        
        self._create_heatmap(
            matrix=matrix,
            start_date=start_date,
            end_date=end_date,
            title=f"Outside Temperature Calendar - {device_name}",
            colormap='viridis',
            missing_color='white',
            output_path=output_path,
            value_label='Temperature (°C)'
        )
        
        logger.info(f"Saved outside temperature calendar: {output_path}")
        return str(output_path)
        return str(output_path)
    
    def _create_heatmap(self,
                       matrix: np.ndarray,
                       start_date: datetime.date,
                       end_date: datetime.date,
                       title: str,
                       colormap: str,
                       missing_color: str,
                       output_path: Path,
                       value_label: str,
                       is_binary: bool = False):
        """Create and save a calendar heatmap image."""
        
        num_days, num_samples = matrix.shape
        
        # Create figure with appropriate size
        fig_width = 14
        fig_height = max(8, num_days * 0.1)  # Scale height with number of days
        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        
        # Create custom colormap with missing value color
        if colormap == 'heating':
            # Custom colormap: white for 0 (no heating), red for 1 (heating active)
            colors = ['white', 'red']
            cmap = mcolors.ListedColormap(colors)
        else:
            cmap = plt.get_cmap(colormap)
            if is_binary:
                # For other binary data, use discrete colors
                cmap = plt.get_cmap(colormap, 2)
        
        # Set color for missing values (NaN)
        cmap.set_bad(color=missing_color)
        
        # Create the heatmap
        if is_binary:
            im = ax.imshow(matrix, aspect='auto', cmap=cmap, 
                          interpolation='nearest', vmin=0, vmax=1)
        else:
            im = ax.imshow(matrix, aspect='auto', cmap=cmap, 
                          interpolation='nearest')
        
        # Set up x-axis (time of day)
        # Show labels every 4 hours (48 samples = 4 hours)
        num_time_labels = 7  # 00:00, 04:00, 08:00, 12:00, 16:00, 20:00, 24:00
        time_tick_indices = np.linspace(0, num_samples - 1, num_time_labels, dtype=int)
        time_labels = [f"{(i * SAMPLE_INTERVAL_MINUTES) // 60:02d}:00" 
                      for i in time_tick_indices]
        ax.set_xticks(time_tick_indices)
        ax.set_xticklabels(time_labels)
        ax.set_xlabel('Time of Day')
        
        # Set up y-axis (dates)
        # Show date labels for reasonable intervals
        if num_days <= 31:
            # Show all days for one month or less
            date_tick_interval = 1
        elif num_days <= 90:
            # Show every 3 days for up to 3 months
            date_tick_interval = 3
        elif num_days <= 180:
            # Show weekly for up to 6 months
            date_tick_interval = 7
        else:
            # Show every 2 weeks for longer periods
            date_tick_interval = 14
        
        date_tick_indices = list(range(0, num_days, date_tick_interval))
        if date_tick_indices[-1] != num_days - 1:
            date_tick_indices.append(num_days - 1)
        
        date_labels = [(start_date + timedelta(days=i)).strftime('%Y-%m-%d') 
                      for i in date_tick_indices]
        ax.set_yticks(date_tick_indices)
        ax.set_yticklabels(date_labels)
        ax.set_ylabel('Date')
        
        # Invert y-axis so earliest date is at top
        ax.invert_yaxis()
        
        # Add colorbar
        if is_binary:
            cbar = plt.colorbar(im, ax=ax, ticks=[0, 1])
            cbar.ax.set_yticklabels(['Off', 'On'])
        else:
            cbar = plt.colorbar(im, ax=ax)
        cbar.set_label(value_label)
        
        # Set title
        ax.set_title(title, fontsize=14, pad=20)
        
        # Tight layout
        plt.tight_layout()
        
        # Save the figure
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
    
    def generate_all_calendars(self):
        """Generate all calendar images for all zones."""
        print("Calendar Image Generation")
        print("=" * 50)
        print()
        
        for device_name in ZONE_DEVICES.keys():
            zone_id = device_name.split('_')[-1]
            
            try:
                # Generate temperature calendar
                print(f"Generating calendars for {device_name} ({ZONE_DEVICES[device_name]})...")
                temp_path = self.generate_temperature_calendar(device_name)
                print(f"✓ Temperature calendar: TempCal_{zone_id}.png")
                
                # Generate heating calendar
                heating_path = self.generate_heating_calendar(device_name)
                print(f"✓ Heating calendar: Heating_{zone_id}.png")
                print()
                
            except Exception as e:
                logger.error(f"Error generating calendars for {device_name}: {e}")
                print(f"✗ Error for {device_name}: {e}")
                print()
        
        # Generate temperature difference calendar
        try:
            print(f"Generating temperature difference calendar...")
            print(f"Devices: {TEMP_DIFF_DEVICES[0]} - {TEMP_DIFF_DEVICES[1]}")
            diff_path = self.generate_temperature_difference_calendar()
            print(f"✓ Temperature difference calendar: TempDiff.png")
            print()
        except Exception as e:
            logger.error(f"Error generating temperature difference calendar: {e}")
            print(f"✗ Error for temperature difference: {e}")
            print()
        
        # Generate outside temperature calendar
        try:
            print(f"Generating outside temperature calendar...")
            print(f"Device: {OUTSIDE_DEVICE}")
            outside_path = self.generate_outside_temperature_calendar()
            print(f"✓ Outside temperature calendar: Temp_Outside.png")
            print()
        except Exception as e:
            logger.error(f"Error generating outside temperature calendar: {e}")
            print(f"✗ Error for outside temperature: {e}")
            print()
        
        print(f"All calendars saved in: {self.output_dir}")


def main():
    """Main function to generate all calendar images."""
    try:
        generator = CalendarImageGenerator()
        generator.generate_all_calendars()
        
    except FileNotFoundError as e:
        print(f"ERROR: Required file not found!")
        print(f"{e}")
        print("\nPlease ensure both temperature_database.json and heating_cycles.json exist.")
        print("Run main.py and detect_heating.py first to generate these files.")
    except Exception as e:
        print(f"ERROR: {e}")
        logger.error(f"Calendar generation error: {e}")


if __name__ == "__main__":
    main()
