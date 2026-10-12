"""
Simple Temperature Data Visualizer

Creates basic, clear visualizations for temperature monitoring data based on
the JSON database. Focuses on core functionality and readability.
"""

import json
import matplotlib.pyplot as plt
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import logging

logger = logging.getLogger(__name__)

# Set up clean plotting style
plt.style.use('default')
plt.rcParams['figure.figsize'] = [12, 8]
plt.rcParams['font.size'] = 10


class SimpleTemperatureVisualizer:
    """Simple visualization generator for temperature monitoring data."""
    
    def __init__(self, json_db_path: str = "data/temperature_database.json", 
                 output_dir: str = "output"):
        self.json_db_path = Path(json_db_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        self.database = self._load_database()
        self._device_dataframes = {}  # Cache for processed device data
        self._summary_cache = None    # Cache for summary data
        
    def preload_all_devices(self):
        """Preload all device data into memory for faster processing."""
        print("Preloading device data...")
        device_names = list(self.database.get('devices', {}).keys())
        
        for i, device_name in enumerate(device_names):
            self._get_device_dataframe(device_name)
            if (i + 1) % 5 == 0 or i == len(device_names) - 1:
                print(f"  Loaded {i + 1}/{len(device_names)} devices...")
        
        print("Preloading complete!")
    
    def create_device_timeline_fast(self, device_name: str, days_limit: Optional[int] = None, 
                                   sample_rate: int = 1) -> str:
        """
        Create a timeline visualization for a single device with optional data sampling.
        
        Args:
            device_name: Name of the device
            days_limit: Limit data to last N days (None for all data)
            sample_rate: Use every Nth data point (1 = all data, 10 = every 10th point)
            
        Returns:
            Path to saved plot
        """
        df = self._get_device_dataframe(device_name)
        
        if df.empty:
            raise ValueError(f"No data found for device '{device_name}'")
        
        # Apply date limit if specified
        if days_limit:
            cutoff_date = df['timestamp'].max() - timedelta(days=days_limit)
            df = df[df['timestamp'] >= cutoff_date]
        
        # Apply sampling for large datasets
        if sample_rate > 1:
            df = df.iloc[::sample_rate]
            suffix_sampling = f"_sampled_{sample_rate}"
        else:
            suffix_sampling = ""
        
        # Create the plot
        fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(14, 10), sharex=True)
        
        # Temperature plot
        ax1.plot(df['timestamp'], df['temperature'], 'b-', linewidth=1.2, label='Temperature')
        ax1.set_ylabel('Temperature (°C)')
        title = f'Temperature Timeline - {device_name}'
        if sample_rate > 1:
            title += f' (Sampled 1:{sample_rate})'
        ax1.set_title(title)
        ax1.grid(True, alpha=0.3)
        ax1.legend()
        
        # Humidity plot
        ax2.plot(df['timestamp'], df['humidity'], 'g-', linewidth=1.2, label='Humidity')
        ax2.set_ylabel('Humidity (%)')
        ax2.set_title(f'Humidity Timeline - {device_name}')
        ax2.grid(True, alpha=0.3)
        ax2.legend()
        
        # Battery plot
        ax3.plot(df['timestamp'], df['battery_mv'], 'r-', linewidth=1.2, label='Battery')
        ax3.set_ylabel('Battery (mV)')
        ax3.set_xlabel('Time')
        ax3.set_title(f'Battery Level - {device_name}')
        ax3.grid(True, alpha=0.3)
        ax3.legend()
        
        # Format x-axis
        fig.autofmt_xdate()
        plt.tight_layout()
        
        # Save the plot
        suffix = f"_last_{days_limit}days" if days_limit else "_full"
        save_path = self.output_dir / f"{device_name}_timeline{suffix}{suffix_sampling}.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved timeline plot: {save_path}")
        return str(save_path)
        
    def _load_database(self) -> Dict:
        """Load the JSON database."""
        if not self.json_db_path.exists():
            raise FileNotFoundError(f"Database not found: {self.json_db_path}")
        
        with open(self.json_db_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _get_device_dataframe(self, device_name: str) -> pd.DataFrame:
        """Get cached DataFrame for a device, creating it if needed."""
        if device_name not in self._device_dataframes:
            if device_name not in self.database.get('devices', {}):
                raise ValueError(f"Device '{device_name}' not found in database")
            
            records = self.database['devices'][device_name].get('records', [])
            if not records:
                return pd.DataFrame()
            
            # Convert all records to DataFrame in one go
            df_data = []
            for record in records:
                df_data.append({
                    'timestamp': pd.to_datetime(record['timestamp']),
                    'temperature': record['temperature'],
                    'humidity': record['humidity'],
                    'battery_mv': record['battery_mv']
                })
            
            df = pd.DataFrame(df_data)
            df = df.sort_values('timestamp')
            self._device_dataframes[device_name] = df
        
        return self._device_dataframes[device_name]
    
    def get_device_data_summary(self) -> Dict[str, Dict]:
        """Get a summary of data available for each device."""
        if self._summary_cache is not None:
            return self._summary_cache
        
        summary = {}
        
        for device_name in self.database.get('devices', {}):
            df = self._get_device_dataframe(device_name)
            
            if df.empty:
                continue
            
            summary[device_name] = {
                'record_count': len(df),
                'first_reading': df['timestamp'].min(),
                'last_reading': df['timestamp'].max(),
                'temperature_range': (df['temperature'].min(), df['temperature'].max()),
                'temperature_avg': df['temperature'].mean(),
                'humidity_range': (df['humidity'].min(), df['humidity'].max()), 
                'humidity_avg': df['humidity'].mean(),
                'battery_range': (df['battery_mv'].min(), df['battery_mv'].max()),
                'battery_avg': df['battery_mv'].mean()
            }
        
        self._summary_cache = summary
        return summary
    
    def create_device_timeline(self, device_name: str, days_limit: Optional[int] = None) -> str:
        """
        Create a timeline visualization for a single device.
        
        Args:
            device_name: Name of the device
            days_limit: Limit data to last N days (None for all data)
            
        Returns:
            Path to saved plot
        """
        df = self._get_device_dataframe(device_name)
        
        if df.empty:
            raise ValueError(f"No data found for device '{device_name}'")
        
        # Apply date limit if specified
        if days_limit:
            cutoff_date = df['timestamp'].max() - timedelta(days=days_limit)
            df = df[df['timestamp'] >= cutoff_date]
        
        # Create the plot
        fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(14, 10), sharex=True)
        
        # Temperature plot
        ax1.plot(df['timestamp'], df['temperature'], 'b-', linewidth=1.2, label='Temperature')
        ax1.set_ylabel('Temperature (°C)')
        ax1.set_title(f'Temperature Timeline - {device_name}')
        ax1.grid(True, alpha=0.3)
        ax1.legend()
        
        # Humidity plot
        ax2.plot(df['timestamp'], df['humidity'], 'g-', linewidth=1.2, label='Humidity')
        ax2.set_ylabel('Humidity (%)')
        ax2.set_title(f'Humidity Timeline - {device_name}')
        ax2.grid(True, alpha=0.3)
        ax2.legend()
        
        # Battery plot
        ax3.plot(df['timestamp'], df['battery_mv'], 'r-', linewidth=1.2, label='Battery')
        ax3.set_ylabel('Battery (mV)')
        ax3.set_xlabel('Time')
        ax3.set_title(f'Battery Level - {device_name}')
        ax3.grid(True, alpha=0.3)
        ax3.legend()
        
        # Format x-axis
        fig.autofmt_xdate()
        plt.tight_layout()
        
        # Save the plot
        suffix = f"_last_{days_limit}days" if days_limit else "_full"
        save_path = self.output_dir / f"{device_name}_timeline{suffix}.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved timeline plot: {save_path}")
        return str(save_path)
    
    def create_temperature_comparison(self, device_names: Optional[List[str]] = None,
                                    days_limit: Optional[int] = None) -> str:
        """
        Create a temperature comparison plot for multiple devices.
        
        Args:
            device_names: List of device names (None for all devices)
            days_limit: Limit data to last N days (None for all data)
            
        Returns:
            Path to saved plot
        """
        if device_names is None:
            device_names = list(self.database.get('devices', {}).keys())
        
        if not device_names:
            raise ValueError("No devices specified or found in database")
        
        fig, ax = plt.subplots(figsize=(14, 8))
        
        # Use simple color cycling
        colors = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'pink', 'gray', 'olive', 'cyan'] * 2
        
        for i, device_name in enumerate(device_names):
            try:
                df = self._get_device_dataframe(device_name)
                if df.empty:
                    logger.warning(f"No data for device '{device_name}', skipping")
                    continue
                
                # Apply date limit if specified
                if days_limit:
                    cutoff_date = df['timestamp'].max() - timedelta(days=days_limit)
                    df = df[df['timestamp'] >= cutoff_date]
                
                if len(df) > 0:
                    ax.plot(df['timestamp'], df['temperature'], 
                           linewidth=1.5, label=device_name, color=colors[i], alpha=0.8)
            except ValueError:
                logger.warning(f"Device '{device_name}' not found, skipping")
                continue
        
        ax.set_xlabel('Time')
        ax.set_ylabel('Temperature (°C)')
        title = f'Temperature Comparison - All Devices'
        if days_limit:
            title += f' (Last {days_limit} days)'
        ax.set_title(title)
        ax.grid(True, alpha=0.3)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # Format x-axis
        fig.autofmt_xdate()
        plt.tight_layout()
        
        # Save the plot
        suffix = f"_last_{days_limit}days" if days_limit else "_full"
        save_path = self.output_dir / f"temperature_comparison{suffix}.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved temperature comparison: {save_path}")
        return str(save_path)
    
    def create_device_summary_chart(self) -> str:
        """
        Create a summary chart showing key statistics for all devices.
        
        Returns:
            Path to saved plot
        """
        summary = self.get_device_data_summary()
        
        if not summary:
            raise ValueError("No device data available for summary")
        
        # Prepare data for plotting
        device_names = list(summary.keys())
        avg_temps = [summary[dev]['temperature_avg'] for dev in device_names]
        avg_humidity = [summary[dev]['humidity_avg'] for dev in device_names]
        record_counts = [summary[dev]['record_count'] for dev in device_names]
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
        
        # Average temperature by device
        bars1 = ax1.bar(device_names, avg_temps, color='lightcoral', alpha=0.7)
        ax1.set_title('Average Temperature by Device')
        ax1.set_ylabel('Temperature (°C)')
        ax1.tick_params(axis='x', rotation=45)
        ax1.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for bar, temp in zip(bars1, avg_temps):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{temp:.1f}°C', ha='center', va='bottom')
        
        # Average humidity by device
        bars2 = ax2.bar(device_names, avg_humidity, color='lightblue', alpha=0.7)
        ax2.set_title('Average Humidity by Device')
        ax2.set_ylabel('Humidity (%)')
        ax2.tick_params(axis='x', rotation=45)
        ax2.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for bar, hum in zip(bars2, avg_humidity):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{hum:.1f}%', ha='center', va='bottom')
        
        # Record count by device
        bars3 = ax3.bar(device_names, record_counts, color='lightgreen', alpha=0.7)
        ax3.set_title('Number of Records by Device')
        ax3.set_ylabel('Record Count')
        ax3.tick_params(axis='x', rotation=45)
        ax3.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for bar, count in zip(bars3, record_counts):
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                    f'{count:,}', ha='center', va='bottom')
        
        # Date range information
        ax4.axis('off')
        info_text = "Data Summary:\n\n"
        
        for device_name in device_names[:5]:  # Show first 5 devices to avoid clutter
            dev_summary = summary[device_name]
            info_text += f"{device_name}:\n"
            info_text += f"  • Records: {dev_summary['record_count']:,}\n"
            info_text += f"  • First: {dev_summary['first_reading'].strftime('%Y-%m-%d %H:%M')}\n"
            info_text += f"  • Last: {dev_summary['last_reading'].strftime('%Y-%m-%d %H:%M')}\n"
            info_text += f"  • Temp Range: {dev_summary['temperature_range'][0]:.1f} - {dev_summary['temperature_range'][1]:.1f}°C\n\n"
        
        if len(device_names) > 5:
            info_text += f"... and {len(device_names) - 5} more devices"
        
        ax4.text(0.05, 0.95, info_text, transform=ax4.transAxes, fontsize=9,
                verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.8))
        
        plt.tight_layout()
        
        save_path = self.output_dir / "device_summary_chart.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved device summary chart: {save_path}")
        return str(save_path)
    
    def create_quick_overview(self) -> str:
        """
        Create a quick overview visualization with key insights.
        
        Returns:
            Path to saved plot
        """
        summary = self.get_device_data_summary()
        
        if not summary:
            raise ValueError("No device data available")
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 10))
        fig.suptitle('Temperature Monitoring - Quick Overview', fontsize=16, fontweight='bold')
        
        device_names = list(summary.keys())
        
        # 1. Current status (latest temperatures)
        latest_temps = []
        latest_humidity = []
        
        for device_name in device_names:
            df = self._get_device_dataframe(device_name)
            if not df.empty:
                # Get the most recent record
                latest_idx = df['timestamp'].idxmax()
                latest_temps.append(df.loc[latest_idx, 'temperature'])
                latest_humidity.append(df.loc[latest_idx, 'humidity'])
        
        if latest_temps:
            bars1 = ax1.bar(device_names, latest_temps, color='orange', alpha=0.7)
            ax1.set_title('Latest Temperature Readings')
            ax1.set_ylabel('Temperature (°C)')
            ax1.tick_params(axis='x', rotation=45)
            ax1.grid(True, alpha=0.3)
            
            for bar, temp in zip(bars1, latest_temps):
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{temp:.1f}°C', ha='center', va='bottom', fontsize=8)
        
        # 2. Temperature ranges (min-max)
        temp_mins = [summary[dev]['temperature_range'][0] for dev in device_names]
        temp_maxs = [summary[dev]['temperature_range'][1] for dev in device_names]
        temp_ranges = [max_t - min_t for min_t, max_t in zip(temp_mins, temp_maxs)]
        
        bars2 = ax2.bar(device_names, temp_ranges, color='red', alpha=0.6)
        ax2.set_title('Temperature Variation Range')
        ax2.set_ylabel('Temperature Range (°C)')
        ax2.tick_params(axis='x', rotation=45)
        ax2.grid(True, alpha=0.3)
        
        for bar, range_val in zip(bars2, temp_ranges):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.05,
                    f'{range_val:.1f}°C', ha='center', va='bottom', fontsize=8)
        
        # 3. Data availability timeline
        first_dates = [summary[dev]['first_reading'] for dev in device_names]
        last_dates = [summary[dev]['last_reading'] for dev in device_names]
        
        for i, (device, first, last) in enumerate(zip(device_names, first_dates, last_dates)):
            ax3.barh(i, (last - first).days, left=first, height=0.6, 
                    alpha=0.7, label=device if i < 10 else "")  # Limit legend entries
        
        ax3.set_title('Data Collection Timeline')
        ax3.set_xlabel('Date')
        ax3.set_yticks(range(len(device_names)))
        ax3.set_yticklabels(device_names, fontsize=8)
        ax3.grid(True, alpha=0.3)
        
        # 4. Summary statistics
        ax4.axis('off')
        
        total_records = sum(summary[dev]['record_count'] for dev in device_names)
        avg_temp_all = sum(summary[dev]['temperature_avg'] for dev in device_names) / len(device_names)
        avg_humidity_all = sum(summary[dev]['humidity_avg'] for dev in device_names) / len(device_names)
        
        oldest_reading = min(summary[dev]['first_reading'] for dev in device_names)
        newest_reading = max(summary[dev]['last_reading'] for dev in device_names)
        
        stats_text = f"""SUMMARY STATISTICS

* Total Devices: {len(device_names)}
* Total Records: {total_records:,}
* Data Period: {oldest_reading.strftime('%Y-%m-%d')} to {newest_reading.strftime('%Y-%m-%d')}
* Average Temperature: {avg_temp_all:.1f}°C
* Average Humidity: {avg_humidity_all:.1f}%

TOP DEVICES BY RECORDS:
"""
        
        # Add top 5 devices by record count
        sorted_devices = sorted(device_names, key=lambda x: summary[x]['record_count'], reverse=True)
        for i, device in enumerate(sorted_devices[:5]):
            stats_text += f"{i+1}. {device}: {summary[device]['record_count']:,} records\n"
        
        ax4.text(0.05, 0.95, stats_text, transform=ax4.transAxes, fontsize=10,
                verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='lightgray', alpha=0.8))
        
        plt.tight_layout()
        
        save_path = self.output_dir / "quick_overview.png"
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved quick overview: {save_path}")
        return str(save_path)


def main():
    """Main function to demonstrate the simple visualizer."""
    try:
        visualizer = SimpleTemperatureVisualizer()
        
        print("Simple Temperature Monitoring Visualizations")
        print("=" * 50)
        
        # Preload all device data for faster processing
        visualizer.preload_all_devices()
        
        # Get summary of available data
        summary = visualizer.get_device_data_summary()
        device_names = list(summary.keys())
        
        if not device_names:
            print("ERROR: No device data found in database!")
            return
        
        print(f"Found data for {len(device_names)} devices:")
        for device in device_names:
            records = summary[device]['record_count']
            first = summary[device]['first_reading'].strftime('%Y-%m-%d')
            last = summary[device]['last_reading'].strftime('%Y-%m-%d')
            print(f"  • {device}: {records:,} records ({first} to {last})")
        
        print("\nGenerating visualizations...")
        
        # Generate overview
        overview_path = visualizer.create_quick_overview()
        print(f"✓ Quick overview: {Path(overview_path).name}")
        
        # Generate summary chart
        summary_path = visualizer.create_device_summary_chart()
        print(f"✓ Device summary: {Path(summary_path).name}")
        
        # Generate temperature comparison
        comparison_path = visualizer.create_temperature_comparison()
        print(f"✓ Temperature comparison: {Path(comparison_path).name}")
        
        # Generate individual device timelines for devices with enough data
        for device_name in device_names:
            if summary[device_name]['record_count'] > 50:  # Only for devices with sufficient data
                try:
                    timeline_path = visualizer.create_device_timeline(device_name)
                    print(f"✓ {device_name} timeline: {Path(timeline_path).name}")
                except Exception as e:
                    print(f"WARNING: Could not create timeline for {device_name}: {e}")
        
        print(f"\nAll visualizations saved in: {visualizer.output_dir}")
        
    except FileNotFoundError:
        print("ERROR: Temperature database not found!")
        print("Please run the data importer first to create the database.")
    except Exception as e:
        print(f"ERROR: Error creating visualizations: {e}")
        logger.error(f"Visualization error: {e}")


if __name__ == "__main__":
    main()
