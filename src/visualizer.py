"""
Data visualization module for temperature monitoring data.
"""

import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
from typing import List, Dict, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class TemperatureVisualizer:
    """Class for creating temperature data visualizations."""
    
    def __init__(self, output_dir: str = "output"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        # Set up matplotlib style
        plt.style.use('seaborn-v0_8')
        sns.set_palette("husl")
    
    def create_temperature_timeline(self, device_data: Dict, save_path: Optional[str] = None) -> str:
        """
        Create a temperature timeline plot for a single device.
        
        Args:
            device_data: Dictionary containing device data
            save_path: Optional custom save path
            
        Returns:
            Path to the saved plot
        """
        device_name = device_data['device_name']
        data = device_data['data']
        
        if not data:
            raise ValueError(f"No data available for device {device_name}")
        
        # Convert to DataFrame for easier plotting
        df = pd.DataFrame(data)
        
        # Create the plot
        fig, (ax1, ax2, ax3) = plt.subplots(3, 1, figsize=(12, 10), sharex=True)
        
        # Temperature plot
        ax1.plot(df['timestamp'], df['temperature'], 'b-', linewidth=1.5, label='Temperature')
        ax1.set_ylabel('Temperature (°C)')
        ax1.set_title(f'Temperature Data - {device_name}')
        ax1.grid(True, alpha=0.3)
        ax1.legend()
        
        # Humidity plot
        ax2.plot(df['timestamp'], df['humidity'], 'g-', linewidth=1.5, label='Humidity')
        ax2.set_ylabel('Humidity (%)')
        ax2.grid(True, alpha=0.3)
        ax2.legend()
        
        # Battery plot
        ax3.plot(df['timestamp'], df['battery_mv'], 'r-', linewidth=1.5, label='Battery')
        ax3.set_ylabel('Battery (mV)')
        ax3.set_xlabel('Time')
        ax3.grid(True, alpha=0.3)
        ax3.legend()
        
        # Format x-axis
        fig.autofmt_xdate()
        plt.tight_layout()
        
        # Save the plot
        if save_path is None:
            save_path = self.output_dir / f"{device_name}_timeline.png"
        
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved temperature timeline plot: {save_path}")
        return str(save_path)
    
    def create_multi_device_comparison(self, all_device_data: List[Dict], 
                                     metric: str = 'temperature',
                                     save_path: Optional[str] = None) -> str:
        """
        Create a comparison plot for multiple devices.
        
        Args:
            all_device_data: List of device data dictionaries
            metric: Metric to compare ('temperature', 'humidity', 'battery_mv')
            save_path: Optional custom save path
            
        Returns:
            Path to the saved plot
        """
        if not all_device_data:
            raise ValueError("No device data provided")
        
        fig, ax = plt.subplots(figsize=(14, 8))
        
        for device_data in all_device_data:
            device_name = device_data['device_name']
            data = device_data['data']
            
            if not data:
                continue
                
            df = pd.DataFrame(data)
            ax.plot(df['timestamp'], df[metric], linewidth=1.5, 
                   label=device_name, alpha=0.8)
        
        # Customize plot
        ax.set_xlabel('Time')
        
        if metric == 'temperature':
            ax.set_ylabel('Temperature (°C)')
            ax.set_title('Temperature Comparison - All Devices')
        elif metric == 'humidity':
            ax.set_ylabel('Humidity (%)')
            ax.set_title('Humidity Comparison - All Devices')
        elif metric == 'battery_mv':
            ax.set_ylabel('Battery (mV)')
            ax.set_title('Battery Level Comparison - All Devices')
        
        ax.grid(True, alpha=0.3)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # Format x-axis
        fig.autofmt_xdate()
        plt.tight_layout()
        
        # Save the plot
        if save_path is None:
            save_path = self.output_dir / f"multi_device_{metric}_comparison.png"
        
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved multi-device comparison plot: {save_path}")
        return str(save_path)
    
    def create_statistics_heatmap(self, all_device_data: List[Dict], 
                                save_path: Optional[str] = None) -> str:
        """
        Create a heatmap showing statistics for all devices.
        
        Args:
            all_device_data: List of device data dictionaries
            save_path: Optional custom save path
            
        Returns:
            Path to the saved plot
        """
        if not all_device_data:
            raise ValueError("No device data provided")
        
        # Calculate statistics for each device
        stats_data = []
        for device_data in all_device_data:
            device_name = device_data['device_name']
            data = device_data['data']
            
            if not data:
                continue
                
            df = pd.DataFrame(data)
            
            stats = {
                'Device': device_name,
                'Avg Temp (°C)': df['temperature'].mean(),
                'Max Temp (°C)': df['temperature'].max(),
                'Min Temp (°C)': df['temperature'].min(),
                'Avg Humidity (%)': df['humidity'].mean(),
                'Max Humidity (%)': df['humidity'].max(),
                'Min Humidity (%)': df['humidity'].min(),
                'Avg Battery (mV)': df['battery_mv'].mean(),
                'Min Battery (mV)': df['battery_mv'].min(),
                'Records': len(data)
            }
            stats_data.append(stats)
        
        # Create DataFrame and heatmap
        stats_df = pd.DataFrame(stats_data)
        stats_df.set_index('Device', inplace=True)
        
        # Normalize data for better heatmap visualization
        numeric_cols = stats_df.select_dtypes(include=[float, int]).columns
        normalized_df = stats_df[numeric_cols].copy()
        
        fig, ax = plt.subplots(figsize=(12, 8))
        sns.heatmap(normalized_df.T, annot=True, fmt='.2f', cmap='viridis', 
                   ax=ax, cbar_kws={'label': 'Normalized Values'})
        
        ax.set_title('Device Statistics Heatmap')
        ax.set_xlabel('Devices')
        ax.set_ylabel('Metrics')
        
        plt.tight_layout()
        
        # Save the plot
        if save_path is None:
            save_path = self.output_dir / "statistics_heatmap.png"
        
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info(f"Saved statistics heatmap: {save_path}")
        return str(save_path)
