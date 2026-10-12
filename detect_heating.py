#!/usr/bin/env python3
"""
Heating Detection System

Analyzes temperature data to detect heating cycles for two zones:
- Zone 1: T8_Z1 device
- Zone 2: T6_Z2 device

Heating detection logic:
- Heating starts when temperature rises +5°C above the daily minimum
- Heating ends when temperature drops 1°C below the cycle maximum
- Gaps shorter than 15 minutes between cycles are merged into a single cycle
"""

import json
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
TEMP_RISE_ABOVE_MIN = 5.0     # Degrees Celsius above daily minimum to start heating
TEMP_DROP_BELOW_MAX = -1.0    # Degrees Celsius below cycle maximum to end heating
MIN_GAP_MINUTES = 15          # Minimum gap between cycles to consider them separate
ZONE_DEVICES = {
    'T8_Z1': 'Zone 1',
    'T6_Z2': 'Zone 2'
}


class HeatingDetector:
    """Detects heating cycles from temperature monitoring data using daily min/cycle max heuristics."""
    
    def __init__(self, json_db_path: str = "data/temperature_database.json", 
                 output_dir: str = "output"):
        self.json_db_path = Path(json_db_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        self.database = self._load_database()
        
    def _load_database(self) -> Dict:
        """Load the JSON database."""
        if not self.json_db_path.exists():
            raise FileNotFoundError(f"Database not found: {self.json_db_path}")
        
        with open(self.json_db_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _get_device_data(self, device_name: str) -> pd.DataFrame:
        """Get device data as a sorted DataFrame."""
        if device_name not in self.database.get('devices', {}):
            raise ValueError(f"Device '{device_name}' not found in database")
        
        records = self.database['devices'][device_name].get('records', [])
        if not records:
            logger.warning(f"No data found for device '{device_name}'")
            return pd.DataFrame()
        
        # Convert to DataFrame
        df_data = []
        for record in records:
            df_data.append({
                'timestamp': pd.to_datetime(record['timestamp']),
                'temperature': record['temperature']
            })
        
        df = pd.DataFrame(df_data)
        df = df.sort_values('timestamp').reset_index(drop=True)
        
        # Add date column for daily grouping
        df['date'] = df['timestamp'].dt.date
        
        logger.info(f"Loaded {len(df)} records for device {device_name}")
        return df
    
    def _calculate_daily_minimums(self, df: pd.DataFrame) -> Dict[str, float]:
        """Calculate minimum temperature for each day."""
        daily_mins = {}
        
        for date, group in df.groupby('date'):
            if not group.empty:
                daily_mins[str(date)] = group['temperature'].min()
                logger.debug(f"Daily minimum for {date}: {daily_mins[str(date)]:.1f}°C")
        
        return daily_mins
    
    def _detect_raw_cycles(self, df: pd.DataFrame) -> List[Dict]:
        """Detect raw heating cycles using daily minimum and cycle maximum heuristics."""
        if df.empty or len(df) < 2:
            return []
        
        logger.info(f"Processing {len(df)} temperature readings with daily min/cycle max detection")
        
        # Calculate daily minimums
        daily_mins = self._calculate_daily_minimums(df)
        
        cycles = []
        in_cycle = False
        cycle_start = None
        cycle_max_temp = None
        cycle_start_temp = None
        
        for i in range(len(df)):
            current_timestamp = df.iloc[i]['timestamp']
            current_temp = df.iloc[i]['temperature']
            current_date = str(df.iloc[i]['date'])
            
            daily_min = daily_mins.get(current_date)
            if daily_min is None:
                logger.warning(f"No daily minimum found for {current_date}, skipping")
                continue
            
            if not in_cycle:
                # Check for heating start: temp >= daily_min + 5°C
                heating_threshold = daily_min + TEMP_RISE_ABOVE_MIN
                if current_temp >= heating_threshold:
                    cycle_start = current_timestamp
                    cycle_start_temp = current_temp
                    cycle_max_temp = current_temp
                    in_cycle = True
                    logger.debug(f"Heating start at {current_timestamp}: {current_temp:.1f}°C "
                               f"(daily min: {daily_min:.1f}°C, threshold: {heating_threshold:.1f}°C)")
            else:
                # Update cycle maximum
                if cycle_max_temp is not None and current_temp > cycle_max_temp:
                    cycle_max_temp = current_temp
                    logger.debug(f"New cycle max at {current_timestamp}: {cycle_max_temp:.1f}°C")
                
                # Check for heating end: temp <= cycle_max - 1°C
                if cycle_max_temp is not None:
                    end_threshold = cycle_max_temp + TEMP_DROP_BELOW_MAX
                    if current_temp <= end_threshold:
                        if cycle_start:
                            duration_minutes = (current_timestamp - cycle_start).total_seconds() / 60
                            cycle_data = {
                                'start': cycle_start,
                                'end': current_timestamp,
                                'max_temp': cycle_max_temp,
                                'duration_minutes': duration_minutes
                            }
                            cycles.append(cycle_data)
                            logger.debug(f"Heating end at {current_timestamp}: {current_temp:.1f}°C "
                                       f"(cycle max: {cycle_max_temp:.1f}°C, threshold: {end_threshold:.1f}°C, "
                                       f"duration: {duration_minutes:.1f}min)")
                        
                        in_cycle = False
                        cycle_start = None
                        cycle_max_temp = None
                        cycle_start_temp = None
        
        # Handle case where data ends while in a cycle
        if in_cycle and cycle_start:
            duration_minutes = (df.iloc[-1]['timestamp'] - cycle_start).total_seconds() / 60
            cycle_data = {
                'start': cycle_start,
                'end': df.iloc[-1]['timestamp'],
                'max_temp': cycle_max_temp,
                'duration_minutes': duration_minutes
            }
            cycles.append(cycle_data)
            logger.debug(f"Heating cycle ended at data end: duration {duration_minutes:.1f}min")
        
        logger.info(f"Detected {len(cycles)} raw heating cycles")
        return cycles
    
    def _merge_cycles_with_gaps(self, raw_cycles: List[Dict]) -> List[Dict]:
        """Merge cycles with gaps shorter than MIN_GAP_MINUTES."""
        if not raw_cycles:
            return []
        
        merged_cycles = []
        current_cycle = raw_cycles[0].copy()
        
        for i in range(1, len(raw_cycles)):
            next_cycle = raw_cycles[i]
            gap_duration = (next_cycle['start'] - current_cycle['end']).total_seconds() / 60  # minutes
            
            if gap_duration <= MIN_GAP_MINUTES:
                # Merge cycles - extend current cycle to include the next one
                current_cycle['end'] = next_cycle['end']
                current_cycle['max_temp'] = max(current_cycle['max_temp'], next_cycle['max_temp'])
                current_cycle['duration_minutes'] = (current_cycle['end'] - current_cycle['start']).total_seconds() / 60
                logger.debug(f"Merged cycles with {gap_duration:.1f}min gap, new duration: {current_cycle['duration_minutes']:.1f}min")
            else:
                # Gap is too large, finalize current cycle and start new one
                merged_cycles.append(current_cycle)
                current_cycle = next_cycle.copy()
        
        # Add the final cycle
        merged_cycles.append(current_cycle)
        
        return merged_cycles
    
    def detect_heating_cycles(self, device_name: str) -> List[Dict[str, str]]:
        """
        Detect heating cycles for a specific device.
        
        Returns:
            List of cycles with start/end timestamps, max temp, and duration
        """
        logger.info(f"Detecting heating cycles for {device_name}...")
        
        df = self._get_device_data(device_name)
        if df.empty:
            return []
        
        # Step 1: Detect raw cycles
        raw_cycles = self._detect_raw_cycles(df)
        logger.info(f"Found {len(raw_cycles)} raw cycles for {device_name}")
        
        # Step 2: Merge cycles with short gaps
        merged_cycles = self._merge_cycles_with_gaps(raw_cycles)
        logger.info(f"After merging gaps <= {MIN_GAP_MINUTES}min: {len(merged_cycles)} cycles for {device_name}")
        
        # Step 3: Convert to required format
        cycles = []
        for cycle in merged_cycles:
            cycles.append({
                'start': cycle['start'].isoformat(),
                'end': cycle['end'].isoformat(),
                'maxtemp': f"{cycle['max_temp']:.1f}",
                'durationMinutes': f"{cycle['duration_minutes']:.0f}"
            })
        
        return cycles
    
    def analyze_all_zones(self) -> Dict[str, List[Dict[str, str]]]:
        """Analyze heating cycles for all configured zones."""
        all_cycles = {}
        
        for device_name in ZONE_DEVICES.keys():
            try:
                cycles = self.detect_heating_cycles(device_name)
                all_cycles[device_name] = cycles
                
                if cycles:
                    total_duration = sum(float(cycle['durationMinutes']) for cycle in cycles) / 60  # Convert to hours
                    avg_max_temp = sum(float(cycle['maxtemp']) for cycle in cycles) / len(cycles)
                    logger.info(f"{device_name}: {len(cycles)} cycles, total {total_duration:.1f} hours, avg max temp {avg_max_temp:.1f}°C")
                else:
                    logger.info(f"{device_name}: No heating cycles detected")
                    
            except Exception as e:
                logger.error(f"Error analyzing {device_name}: {e}")
                all_cycles[device_name] = []
        
        return all_cycles
    
    def save_cycles_json(self, cycles_data: Dict[str, List[Dict[str, str]]]) -> str:
        """Save heating cycles to JSON file."""
        output_path = self.output_dir / "heating_cycles.json"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(cycles_data, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Saved heating cycles to: {output_path}")
        return str(output_path)
    
    def _get_full_date_range(self, cycles: List[Dict[str, str]]) -> pd.DatetimeIndex:
        """Get complete date range from first to last cycle date."""
        if not cycles:
            return pd.DatetimeIndex([])
        
        dates = [pd.to_datetime(cycle['start']).date() for cycle in cycles]
        min_date = min(dates)
        max_date = max(dates)
        
        return pd.date_range(start=min_date, end=max_date, freq='D')
    
    def _calculate_cycles_per_day(self, cycles: List[Dict[str, str]]) -> Dict[str, int]:
        """Calculate number of cycles per day."""
        cycles_per_day = {}
        
        for cycle in cycles:
            start_time = pd.to_datetime(cycle['start'])
            # Assign cycle to the day it started (as per specification)
            day_key = start_time.strftime('%Y-%m-%d')
            cycles_per_day[day_key] = cycles_per_day.get(day_key, 0) + 1
        
        return cycles_per_day
    
    def _calculate_cycles_per_day_complete(self, cycles: List[Dict[str, str]]) -> pd.DataFrame:
        """Calculate number of cycles per day with complete date range (including zeros)."""
        if not cycles:
            return pd.DataFrame(columns=['date', 'cycles'])
        
        # Get basic cycles per day
        cycles_per_day = self._calculate_cycles_per_day(cycles)
        
        # Get complete date range
        date_range = self._get_full_date_range(cycles)
        
        # Create complete DataFrame with zeros for missing days
        complete_data = []
        for date in date_range:
            day_key = date.strftime('%Y-%m-%d')
            cycle_count = cycles_per_day.get(day_key, 0)
            complete_data.append({
                'date': date,
                'cycles': cycle_count
            })
        
        return pd.DataFrame(complete_data)
    
    def _calculate_duration_per_day(self, cycles: List[Dict[str, str]]) -> Dict[str, float]:
        """Calculate total heating duration per day in hours."""
        duration_per_day = {}
        
        for cycle in cycles:
            start_time = pd.to_datetime(cycle['start'])
            duration_hours = float(cycle['durationMinutes']) / 60.0
            
            # Assign cycle to the day it started (as per specification)
            day_key = start_time.strftime('%Y-%m-%d')
            duration_per_day[day_key] = duration_per_day.get(day_key, 0) + duration_hours
        
        return duration_per_day
        
    def _calculate_duration_per_day_complete(self, cycles: List[Dict[str, str]]) -> pd.DataFrame:
        """Calculate total heating duration per day with complete date range (including zeros)."""
        if not cycles:
            return pd.DataFrame(columns=['date', 'duration'])
        
        # Get basic duration per day
        duration_per_day = self._calculate_duration_per_day(cycles)
        
        # Get complete date range
        date_range = self._get_full_date_range(cycles)
        
        # Create complete DataFrame with zeros for missing days
        complete_data = []
        for date in date_range:
            day_key = date.strftime('%Y-%m-%d')
            duration_hours = duration_per_day.get(day_key, 0.0)
            complete_data.append({
                'date': date,
                'duration': duration_hours
            })
        
        return pd.DataFrame(complete_data)
    
    def create_daily_cycle_chart(self, device_name: str, cycles: List[Dict[str, str]]) -> str:
        """Create a chart showing cycles per day for a device."""
        df_plot = self._calculate_cycles_per_day_complete(cycles)
        
        if df_plot.empty:
            logger.warning(f"No cycles to plot for {device_name}")
            return ""
        
        # Create the plot
        plt.figure(figsize=(14, 6))
        plt.plot(df_plot['date'], df_plot['cycles'], 'b.', markersize=5)
        
        plt.title(f'Heating Cycles Per Day - {device_name} ({ZONE_DEVICES.get(device_name, "Unknown Zone")})')
        plt.xlabel('Date')
        plt.ylabel('Number of Cycles')
        plt.grid(True, alpha=0.3)
        
        # Format x-axis
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_locator(mdates.WeekdayLocator(interval=7))  # Weekly ticks
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        
        # Set y-axis to start from 0
        plt.ylim(bottom=0)
        
        plt.tight_layout()
        
        # Save the plot
        zone_id = device_name.split('_')[-1]  # Extract Z1 or Z2
        save_path = self.output_dir / f"heating_cycle_per_day_{zone_id}.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        # Save CSV data
        csv_path = self.output_dir / f"heating_cycle_per_day_{zone_id}.csv"
        df_plot.to_csv(csv_path, index=False)
        
        logger.info(f"Saved daily cycles chart: {save_path}")
        logger.info(f"Saved daily cycles data: {csv_path}")
        return str(save_path)
    
    def create_daily_duration_chart(self, device_name: str, cycles: List[Dict[str, str]]) -> str:
        """Create a chart showing heating duration per day for a device."""
        df_plot = self._calculate_duration_per_day_complete(cycles)
        
        if df_plot.empty:
            logger.warning(f"No duration data to plot for {device_name}")
            return ""
        
        # Create the plot
        plt.figure(figsize=(14, 6))
        plt.plot(df_plot['date'], df_plot['duration'], 'r.', markersize=5)
        
        plt.title(f'Heating Duration Per Day - {device_name} ({ZONE_DEVICES.get(device_name, "Unknown Zone")})')
        plt.xlabel('Date')
        plt.ylabel('Heating Duration (hours)')
        plt.grid(True, alpha=0.3)
        
        # Format x-axis
        plt.xticks(rotation=45)
        plt.gca().xaxis.set_major_locator(mdates.WeekdayLocator(interval=7))  # Weekly ticks
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        
        # Set y-axis to start from 0
        plt.ylim(bottom=0)
        
        plt.tight_layout()
        
        # Save the plot
        zone_id = device_name.split('_')[-1]  # Extract Z1 or Z2
        save_path = self.output_dir / f"heating_duration_per_day_{zone_id}.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        # Save CSV data
        csv_path = self.output_dir / f"heating_duration_per_day_{zone_id}.csv"
        df_plot.to_csv(csv_path, index=False)
        
        logger.info(f"Saved daily duration chart: {save_path}")
        logger.info(f"Saved daily duration data: {csv_path}")
        return str(save_path)
    
    def generate_summary_report(self, cycles_data: Dict[str, List[Dict[str, str]]]) -> str:
        """Generate a summary report of heating analysis."""
        report_lines = []
        report_lines.append("HEATING DETECTION ANALYSIS SUMMARY")
        report_lines.append("=" * 50)
        report_lines.append(f"Heating start: temperature >= daily_min + {TEMP_RISE_ABOVE_MIN}°C")
        report_lines.append(f"Heating end: temperature <= cycle_max {TEMP_DROP_BELOW_MAX}°C")
        report_lines.append(f"Gap merging threshold: <= {MIN_GAP_MINUTES} minutes")
        report_lines.append("")
        
        for device_name, cycles in cycles_data.items():
            zone_name = ZONE_DEVICES.get(device_name, "Unknown Zone")
            report_lines.append(f"{device_name} ({zone_name}):")
            
            if not cycles:
                report_lines.append("  No heating cycles detected")
                report_lines.append("")
                continue
            
            # Calculate statistics
            total_cycles = len(cycles)
            durations_minutes = [float(cycle['durationMinutes']) for cycle in cycles]
            max_temps = [float(cycle['maxtemp']) for cycle in cycles]
            cycles_per_day = self._calculate_cycles_per_day(cycles)
            duration_per_day = self._calculate_duration_per_day(cycles)
            
            total_heating_hours = sum(durations_minutes) / 60.0
            avg_cycle_duration = sum(durations_minutes) / total_cycles if total_cycles > 0 else 0
            avg_max_temp = sum(max_temps) / total_cycles if total_cycles > 0 else 0
            max_cycles_per_day = max(cycles_per_day.values()) if cycles_per_day else 0
            avg_cycles_per_day = sum(cycles_per_day.values()) / len(cycles_per_day) if cycles_per_day else 0
            max_duration_per_day = max(duration_per_day.values()) if duration_per_day else 0
            avg_duration_per_day = sum(duration_per_day.values()) / len(duration_per_day) if duration_per_day else 0
            
            report_lines.append(f"  Total cycles: {total_cycles}")
            report_lines.append(f"  Total heating time: {total_heating_hours:.1f} hours")
            report_lines.append(f"  Average cycle duration: {avg_cycle_duration:.1f} minutes")
            report_lines.append(f"  Average maximum temperature: {avg_max_temp:.1f}°C")
            report_lines.append(f"  Max cycles per day: {max_cycles_per_day}")
            report_lines.append(f"  Average cycles per day: {avg_cycles_per_day:.1f}")
            report_lines.append(f"  Max duration per day: {max_duration_per_day:.1f} hours")
            report_lines.append(f"  Average duration per day: {avg_duration_per_day:.1f} hours")
            report_lines.append(f"  Date range: {min(cycles_per_day.keys())} to {max(cycles_per_day.keys())}")
            report_lines.append("")
        
        # Save report
        report_path = self.output_dir / "heating_analysis_summary.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))
        
        logger.info(f"Saved summary report: {report_path}")
        return str(report_path)


def main():
    """Main function to run heating detection analysis."""
    try:
        detector = HeatingDetector()
        
        print("Heating Detection Analysis")
        print("=" * 40)
        print(f"Heating start: temp >= daily_min + {TEMP_RISE_ABOVE_MIN}°C")
        print(f"Heating end: temp <= cycle_max {TEMP_DROP_BELOW_MAX}°C") 
        print(f"Gap merging: <= {MIN_GAP_MINUTES} minutes")
        print(f"Zones: {', '.join(f'{k} ({v})' for k, v in ZONE_DEVICES.items())}")
        print()
        
        # Analyze all zones
        print("Analyzing heating cycles...")
        cycles_data = detector.analyze_all_zones()
        
        # Save JSON output
        json_path = detector.save_cycles_json(cycles_data)
        print(f"✓ Saved cycles data: {Path(json_path).name}")
        
        # Generate daily cycle charts
        print("\nGenerating daily cycle charts...")
        for device_name, cycles in cycles_data.items():
            if cycles:
                chart_path = detector.create_daily_cycle_chart(device_name, cycles)
                if chart_path:
                    zone_id = device_name.split('_')[-1]
                    print(f"✓ Saved cycle chart: heating_cycle_per_day_{zone_id}.png")
                    print(f"✓ Saved cycle data: heating_cycle_per_day_{zone_id}.csv")
            else:
                print(f"⚠ No cycles to chart for {device_name}")
        
        # Generate daily duration charts
        print("\nGenerating daily duration charts...")
        for device_name, cycles in cycles_data.items():
            if cycles:
                duration_chart_path = detector.create_daily_duration_chart(device_name, cycles)
                if duration_chart_path:
                    zone_id = device_name.split('_')[-1]
                    print(f"✓ Saved duration chart: heating_duration_per_day_{zone_id}.png")
                    print(f"✓ Saved duration data: heating_duration_per_day_{zone_id}.csv")
            else:
                print(f"⚠ No duration data to chart for {device_name}")
        
        # Generate summary report
        report_path = detector.generate_summary_report(cycles_data)
        print(f"✓ Saved summary: {Path(report_path).name}")
        
        print(f"\nAll outputs saved in: {detector.output_dir}")
        
    except FileNotFoundError:
        print("ERROR: Temperature database not found!")
        print("Please ensure data/temperature_database.json exists.")
    except Exception as e:
        print(f"ERROR: {e}")
        logger.error(f"Heating detection error: {e}")


if __name__ == "__main__":
    main()