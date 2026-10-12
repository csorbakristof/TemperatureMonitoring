"""
Excel export module for temperature monitoring data.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, Reference
from typing import List, Dict, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class ExcelExporter:
    """Class for exporting temperature data to Excel files."""
    
    def __init__(self, output_dir: str = "output"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
    
    def export_device_data(self, device_data: Dict, save_path: Optional[str] = None) -> str:
        """
        Export single device data to Excel file.
        
        Args:
            device_data: Dictionary containing device data
            save_path: Optional custom save path
            
        Returns:
            Path to the saved Excel file
        """
        device_name = device_data['device_name']
        data = device_data['data']
        
        if not data:
            raise ValueError(f"No data available for device {device_name}")
        
        # Convert to DataFrame
        df = pd.DataFrame(data)
        
        # Save path
        if save_path is None:
            save_path = self.output_dir / f"{device_name}_data.xlsx"
        
        # Create Excel file with formatting
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            # Write raw data
            df.to_excel(writer, sheet_name='Raw Data', index=False)
            
            # Create statistics sheet
            stats_df = self._calculate_statistics(df)
            stats_df.to_excel(writer, sheet_name='Statistics', index=True)
            
            # Create hourly averages
            hourly_df = self._create_hourly_averages(df)
            hourly_df.to_excel(writer, sheet_name='Hourly Averages', index=False)
            
            # Format the workbook
            self._format_workbook(writer.book, device_name)
        
        logger.info(f"Exported device data to Excel: {save_path}")
        return str(save_path)
    
    def export_all_devices(self, all_device_data: List[Dict], 
                          save_path: Optional[str] = None) -> str:
        """
        Export all device data to a single Excel file with multiple sheets.
        
        Args:
            all_device_data: List of device data dictionaries
            save_path: Optional custom save path
            
        Returns:
            Path to the saved Excel file
        """
        if not all_device_data:
            raise ValueError("No device data provided")
        
        if save_path is None:
            save_path = self.output_dir / "all_devices_data.xlsx"
        
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            # Create summary sheet
            summary_data = []
            for device_data in all_device_data:
                device_name = device_data['device_name']
                data = device_data['data']
                
                if not data:
                    continue
                
                df = pd.DataFrame(data)
                summary = {
                    'Device': device_name,
                    'Records': len(data),
                    'Start Time': df['timestamp'].min(),
                    'End Time': df['timestamp'].max(),
                    'Avg Temperature (°C)': df['temperature'].mean(),
                    'Max Temperature (°C)': df['temperature'].max(),
                    'Min Temperature (°C)': df['temperature'].min(),
                    'Avg Humidity (%)': df['humidity'].mean(),
                    'Max Humidity (%)': df['humidity'].max(),
                    'Min Humidity (%)': df['humidity'].min(),
                    'Avg Battery (mV)': df['battery_mv'].mean(),
                    'Min Battery (mV)': df['battery_mv'].min()
                }
                summary_data.append(summary)
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Create individual sheets for each device
            for device_data in all_device_data:
                device_name = device_data['device_name']
                data = device_data['data']
                
                if not data:
                    continue
                
                df = pd.DataFrame(data)
                sheet_name = device_name[:31]  # Excel sheet name limit
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Format the workbook
            self._format_workbook(writer.book, "Multi-Device Report")
        
        logger.info(f"Exported all device data to Excel: {save_path}")
        return str(save_path)
    
    def _calculate_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """Calculate basic statistics for the data."""
        stats = {
            'Temperature (°C)': {
                'Mean': df['temperature'].mean(),
                'Median': df['temperature'].median(),
                'Std Dev': df['temperature'].std(),
                'Min': df['temperature'].min(),
                'Max': df['temperature'].max(),
                'Range': df['temperature'].max() - df['temperature'].min()
            },
            'Humidity (%)': {
                'Mean': df['humidity'].mean(),
                'Median': df['humidity'].median(),
                'Std Dev': df['humidity'].std(),
                'Min': df['humidity'].min(),
                'Max': df['humidity'].max(),
                'Range': df['humidity'].max() - df['humidity'].min()
            },
            'Battery (mV)': {
                'Mean': df['battery_mv'].mean(),
                'Median': df['battery_mv'].median(),
                'Std Dev': df['battery_mv'].std(),
                'Min': df['battery_mv'].min(),
                'Max': df['battery_mv'].max(),
                'Range': df['battery_mv'].max() - df['battery_mv'].min()
            }
        }
        
        return pd.DataFrame(stats).round(2)
    
    def _create_hourly_averages(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create hourly averages of the data."""
        df_copy = df.copy()
        df_copy.set_index('timestamp', inplace=True)
        
        hourly = df_copy.resample('H').agg({
            'temperature': 'mean',
            'humidity': 'mean',
            'battery_mv': 'mean'
        }).round(2)
        
        hourly.reset_index(inplace=True)
        return hourly
    
    def _format_workbook(self, workbook: Workbook, title: str):
        """Apply formatting to the Excel workbook."""
        # Header style
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center")
        
        for sheet in workbook.worksheets:
            # Format headers
            for cell in sheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
