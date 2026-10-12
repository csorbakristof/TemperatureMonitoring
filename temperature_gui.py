#!/usr/bin/env python3
"""
Temperature Monitoring GUI Application

Interactive GUI for visualizing temperature sensor data with:
- Device selection via checkboxes
- Date range picker
- Multiple data types (temperature, humidity, battery, derived statistics)
- Interactive matplotlib plots
- Export functionality

Usage: python temperature_gui.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import pandas as pd
import json
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path
import logging
from typing import Dict, List, Optional, Any

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TemperatureDatabase:
    """Handler for temperature database operations."""
    
    def __init__(self, db_path: str = "data/temperature_database.json"):
        self.db_path = Path(db_path)
        self.database = None
        self.load_database()
    
    def load_database(self) -> None:
        """Load the temperature database."""
        try:
            if not self.db_path.exists():
                raise FileNotFoundError(f"Database file not found: {self.db_path}")
            
            with open(self.db_path, 'r', encoding='utf-8') as f:
                self.database = json.load(f)
            
            logger.info(f"Database loaded: {len(self.database.get('devices', {}))} devices")
            
        except Exception as e:
            logger.error(f"Failed to load database: {e}")
            self.database = {"devices": {}, "metadata": {}}
    
    def get_devices(self) -> List[str]:
        """Get list of available devices."""
        return list(self.database.get('devices', {}).keys())
    
    def get_device_data(self, device_name: str, start_date: datetime = None, 
                       end_date: datetime = None) -> pd.DataFrame:
        """Get device data as pandas DataFrame with optional date filtering."""
        if device_name not in self.database.get('devices', {}):
            return pd.DataFrame()
        
        device_data = self.database['devices'][device_name]
        records = device_data.get('records', [])
        
        if not records:
            return pd.DataFrame()
        
        # Convert to DataFrame
        df = pd.DataFrame(records)
        df['timestamp'] = pd.to_datetime(df['timestamp'])
        
        # Apply date filtering
        if start_date:
            df = df[df['timestamp'] >= start_date]
        if end_date:
            df = df[df['timestamp'] <= end_date]
        
        # Sort by timestamp
        df = df.sort_values('timestamp').reset_index(drop=True)
        
        return df
    
    def get_date_range(self) -> tuple:
        """Get the overall date range of all devices."""
        all_dates = []
        
        for device_data in self.database.get('devices', {}).values():
            records = device_data.get('records', [])
            for record in records:
                try:
                    dt = datetime.fromisoformat(record['timestamp'])
                    all_dates.append(dt)
                except:
                    continue
        
        if not all_dates:
            # Default range if no data
            end_date = datetime.now()
            start_date = end_date - timedelta(days=30)
            return start_date, end_date
        
        return min(all_dates), max(all_dates)


class TemperatureGUI:
    """Main GUI application for temperature monitoring visualization."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Temperature Monitoring - Interactive Visualization")
        self.root.geometry("1400x900")
        
        # Data handlers
        self.db = TemperatureDatabase()
        
        # GUI variables
        self.device_vars = {}
        self.data_type_var = tk.StringVar(value="temperature")
        self.show_stat002_var = tk.BooleanVar(value=False)
        self.show_room_external_diff_var = tk.BooleanVar(value=False)
        
        # Initialize date range
        self.min_date, self.max_date = self.db.get_date_range()
        
        self.setup_gui()
        self.create_plot_area()
        
        # Load initial data
        self.refresh_plot()
    
    def setup_gui(self):
        """Setup the main GUI layout."""
        # Create main frames
        control_frame = ttk.Frame(self.root)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        
        plot_frame = ttk.Frame(self.root)
        plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Control panel
        self.setup_control_panel(control_frame)
        
        # Plot area
        self.plot_frame = plot_frame
    
    def setup_control_panel(self, parent):
        """Setup the control panel with device selection and options."""
        
        # Title
        title_label = ttk.Label(parent, text="Temperature Monitoring", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Device selection
        device_frame = ttk.LabelFrame(parent, text="Select Devices", padding=10)
        device_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create scrollable frame for devices
        canvas = tk.Canvas(device_frame, height=200)
        scrollbar = ttk.Scrollbar(device_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Add device checkboxes
        devices = self.db.get_devices()
        for device in devices:
            var = tk.BooleanVar(value=False)
            # Pre-select key devices for demonstration
            if device in ['T1_BE', 'T3_Kek', 'T2_Terasz']:
                var.set(True)
            
            self.device_vars[device] = var
            cb = ttk.Checkbutton(scrollable_frame, text=device, variable=var,
                               command=self.on_device_selection_change)
            cb.pack(anchor=tk.W, pady=2)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Select/Deselect all buttons
        button_frame = ttk.Frame(device_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Select All", 
                  command=self.select_all_devices).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Deselect All", 
                  command=self.deselect_all_devices).pack(side=tk.LEFT)
        
        # Data type selection
        data_type_frame = ttk.LabelFrame(parent, text="Data Type", padding=10)
        data_type_frame.pack(fill=tk.X, pady=(0, 20))
        
        data_types = [
            ("Temperature (°C)", "temperature"),
            ("Humidity (%)", "humidity"),
            ("Battery (mV)", "battery_mv"),
            ("All in subplots", "all")
        ]
        
        for text, value in data_types:
            rb = ttk.Radiobutton(data_type_frame, text=text, 
                               variable=self.data_type_var, value=value,
                               command=self.refresh_plot)
            rb.pack(anchor=tk.W, pady=2)
        
        # STAT002 checkbox
        stat002_frame = ttk.LabelFrame(parent, text="Derived Statistics", padding=10)
        stat002_frame.pack(fill=tk.X, pady=(0, 20))
        
        stat002_cb = ttk.Checkbutton(stat002_frame, 
                                   text="Show Temperature Difference (STAT002)\n(Room - Intake)", 
                                   variable=self.show_stat002_var,
                                   command=self.refresh_plot)
        stat002_cb.pack(anchor=tk.W, pady=(0, 5))
        
        room_external_cb = ttk.Checkbutton(stat002_frame, 
                                         text="Show Room vs External Difference\n(Room - External)", 
                                         variable=self.show_room_external_diff_var,
                                         command=self.refresh_plot)
        room_external_cb.pack(anchor=tk.W)
        
        # Date range selection
        date_frame = ttk.LabelFrame(parent, text="Date Range", padding=10)
        date_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(date_frame, text="Start Date:").pack(anchor=tk.W)
        self.start_date = DateEntry(date_frame, width=12, background='darkblue',
                                   foreground='white', borderwidth=2,
                                   mindate=self.min_date.date(),
                                   maxdate=self.max_date.date())
        self.start_date.set_date(self.min_date.date())
        self.start_date.pack(anchor=tk.W, pady=(0, 10))
        self.start_date.bind("<<DateEntrySelected>>", lambda e: self.refresh_plot())
        
        ttk.Label(date_frame, text="End Date:").pack(anchor=tk.W)
        self.end_date = DateEntry(date_frame, width=12, background='darkblue',
                                 foreground='white', borderwidth=2,
                                 mindate=self.min_date.date(),
                                 maxdate=self.max_date.date())
        self.end_date.set_date(self.max_date.date())
        self.end_date.pack(anchor=tk.W, pady=(0, 10))
        self.end_date.bind("<<DateEntrySelected>>", lambda e: self.refresh_plot())
        
        # Quick date range buttons
        quick_frame = ttk.Frame(date_frame)
        quick_frame.pack(fill=tk.X)
        
        ttk.Button(quick_frame, text="Last 7 days", 
                  command=lambda: self.set_quick_range(7)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="Last 30 days", 
                  command=lambda: self.set_quick_range(30)).pack(side=tk.LEFT)
        
        # Control buttons
        button_frame = ttk.LabelFrame(parent, text="Actions", padding=10)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Button(button_frame, text="Refresh Plot", 
                  command=self.refresh_plot).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(button_frame, text="Export Plot", 
                  command=self.export_plot).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(button_frame, text="Export Data", 
                  command=self.export_data).pack(fill=tk.X)
        
        # Status label
        self.status_label = ttk.Label(parent, text="Ready", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X, pady=(20, 0))
    
    def create_plot_area(self):
        """Create the matplotlib plot area."""
        self.fig = Figure(figsize=(12, 8), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        # Add navigation toolbar
        toolbar = NavigationToolbar2Tk(self.canvas, self.plot_frame)
        toolbar.update()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def on_device_selection_change(self):
        """Handle device selection changes."""
        selected_count = sum(1 for var in self.device_vars.values() if var.get())
        self.update_status(f"{selected_count} devices selected")
        self.refresh_plot()
    
    def select_all_devices(self):
        """Select all devices."""
        for var in self.device_vars.values():
            var.set(True)
        self.refresh_plot()
    
    def deselect_all_devices(self):
        """Deselect all devices."""
        for var in self.device_vars.values():
            var.set(False)
        self.refresh_plot()
    
    def set_quick_range(self, days: int):
        """Set a quick date range."""
        end_date = self.max_date.date()
        start_date = (self.max_date - timedelta(days=days)).date()
        
        self.start_date.set_date(start_date)
        self.end_date.set_date(end_date)
        self.refresh_plot()
    
    def update_status(self, message: str):
        """Update the status label."""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def get_selected_devices(self) -> List[str]:
        """Get list of selected devices."""
        return [device for device, var in self.device_vars.items() if var.get()]
    
    def refresh_plot(self):
        """Refresh the plot with current selections."""
        try:
            self.update_status("Loading data...")
            
            # Clear previous plot
            self.fig.clear()
            
            # Get selections
            selected_devices = self.get_selected_devices()
            data_type = self.data_type_var.get()
            show_stat002 = self.show_stat002_var.get()
            show_room_external_diff = self.show_room_external_diff_var.get()
            
            # Get date range
            start_date = datetime.combine(self.start_date.get_date(), datetime.min.time())
            end_date = datetime.combine(self.end_date.get_date(), datetime.max.time())
            
            if not selected_devices and not show_stat002 and not show_room_external_diff:
                self.update_status("No devices or statistics selected")
                self.canvas.draw()
                return
            
            # Determine subplot layout
            subplot_count = 1
            if data_type == "all":
                subplot_count = 3  # temperature, humidity, battery
            derived_stats_count = sum([show_stat002, show_room_external_diff])
            if derived_stats_count > 0:
                subplot_count += derived_stats_count
            
            # Create subplots
            if subplot_count == 1:
                axes = [self.fig.add_subplot(111)]
            else:
                axes = []
                for i in range(subplot_count):
                    axes.append(self.fig.add_subplot(subplot_count, 1, i + 1))
            
            # Plot device data
            axes_index = 0
            if selected_devices:
                if data_type == "all":
                    self.plot_all_data_types(axes[axes_index:axes_index+3], selected_devices, start_date, end_date)
                    axes_index += 3
                else:
                    self.plot_single_data_type(axes[axes_index], selected_devices, data_type, 
                                             start_date, end_date)
                    axes_index += 1
            
            # Plot STAT002 data
            if show_stat002:
                stat002_ax = axes[axes_index]
                self.plot_stat002_data(stat002_ax, start_date, end_date)
                axes_index += 1
            
            # Plot Room vs External difference
            if show_room_external_diff:
                room_external_ax = axes[axes_index]
                self.plot_room_external_diff(room_external_ax, start_date, end_date)
                axes_index += 1
            
            # Adjust layout and draw
            self.fig.tight_layout()
            self.canvas.draw()
            
            self.update_status(f"Plot updated - {len(selected_devices)} devices, "
                             f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            
        except Exception as e:
            logger.error(f"Error refreshing plot: {e}")
            self.update_status(f"Error: {str(e)}")
            messagebox.showerror("Plot Error", f"Failed to refresh plot:\n{str(e)}")
    
    def plot_single_data_type(self, ax, devices: List[str], data_type: str, 
                             start_date: datetime, end_date: datetime):
        """Plot a single data type for selected devices."""
        colors = plt.cm.tab10(np.linspace(0, 1, len(devices)))
        
        for i, device in enumerate(devices):
            df = self.db.get_device_data(device, start_date, end_date)
            
            if df.empty or data_type not in df.columns:
                continue
            
            ax.plot(df['timestamp'], df[data_type], 
                   label=device, color=colors[i], alpha=0.8, linewidth=1.5)
        
        # Formatting
        ax.set_xlabel('Time')
        ax.set_ylabel(self.get_data_type_label(data_type))
        ax.set_title(f'{self.get_data_type_label(data_type)} - Time Series')
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(True, alpha=0.3)
        
        # Format x-axis
        self.format_time_axis(ax)
    
    def plot_all_data_types(self, axes: List, devices: List[str], 
                           start_date: datetime, end_date: datetime):
        """Plot all data types in separate subplots."""
        data_types = ['temperature', 'humidity', 'battery_mv']
        colors = plt.cm.tab10(np.linspace(0, 1, len(devices)))
        
        for ax_idx, data_type in enumerate(data_types):
            ax = axes[ax_idx]
            
            for i, device in enumerate(devices):
                df = self.db.get_device_data(device, start_date, end_date)
                
                if df.empty or data_type not in df.columns:
                    continue
                
                ax.plot(df['timestamp'], df[data_type], 
                       label=device, color=colors[i], alpha=0.8, linewidth=1.5)
            
            # Formatting
            ax.set_ylabel(self.get_data_type_label(data_type))
            ax.set_title(f'{self.get_data_type_label(data_type)}')
            ax.grid(True, alpha=0.3)
            
            if ax_idx == 0:  # Only show legend on first subplot
                ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            
            if ax_idx == len(data_types) - 1:  # Only label x-axis on last subplot
                ax.set_xlabel('Time')
            
            self.format_time_axis(ax)
    
    def plot_stat002_data(self, ax, start_date: datetime, end_date: datetime):
        """Plot STAT002 temperature difference data (calculated on-demand)."""
        # Get data for required devices (Room - Intake difference)
        room_df = self.db.get_device_data('T3_Kek', start_date, end_date)  # Room
        intake_df = self.db.get_device_data('T1_BE', start_date, end_date)  # Intake
        
        if room_df.empty or intake_df.empty:
            ax.text(0.5, 0.5, 'STAT002 data not available\nMissing T3_Kek or T1_BE data', 
                   transform=ax.transAxes, ha='center', va='center',
                   bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
            ax.set_title('Temperature Difference (STAT002)')
            return
        
        # Synchronize the data by timestamp (simple approach)
        # Create a merged dataframe on timestamp with tolerance
        room_df = room_df[['timestamp', 'temperature']].rename(columns={'temperature': 'room_temp'})
        intake_df = intake_df[['timestamp', 'temperature']].rename(columns={'temperature': 'intake_temp'})
        
        # Merge with tolerance (find nearest timestamps within 15 minutes)
        tolerance = pd.Timedelta(minutes=15)
        merged_data = []
        
        for _, room_row in room_df.iterrows():
            room_time = room_row['timestamp']
            room_temp = room_row['room_temp']
            
            # Find closest intake temperature reading within tolerance
            time_diffs = abs(intake_df['timestamp'] - room_time)
            if len(time_diffs) == 0:
                continue
                
            closest_idx = time_diffs.idxmin()
            
            if time_diffs.loc[closest_idx] <= tolerance:
                intake_temp = intake_df.loc[closest_idx, 'intake_temp']
                temp_diff = room_temp - intake_temp
                
                merged_data.append({
                    'timestamp': room_time,
                    'room_temp': room_temp,
                    'intake_temp': intake_temp,
                    'temperature_difference': temp_diff
                })
        
        if not merged_data:
            ax.text(0.5, 0.5, 'No synchronized data available\nfor Room vs Intake comparison', 
                   transform=ax.transAxes, ha='center', va='center',
                   bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
            ax.set_title('Temperature Difference (STAT002)')
            return
        
        # Convert to DataFrame for plotting
        diff_df = pd.DataFrame(merged_data)
        
        # Plot temperature difference
        ax.plot(diff_df['timestamp'], diff_df['temperature_difference'], 
               color='red', alpha=0.8, linewidth=1.5, 
               label='Room - Intake (°C)')
        
        # Add horizontal line at zero
        ax.axhline(y=0, color='black', linestyle='--', alpha=0.5, linewidth=1)
        
        # Color areas above/below zero
        ax.fill_between(diff_df['timestamp'], diff_df['temperature_difference'], 0,
                      where=(diff_df['temperature_difference'] > 0), 
                      color='red', alpha=0.2, label='Room warmer')
        ax.fill_between(diff_df['timestamp'], diff_df['temperature_difference'], 0,
                      where=(diff_df['temperature_difference'] < 0), 
                      color='blue', alpha=0.2, label='Intake warmer')
        
        # Add statistics
        mean_diff = diff_df['temperature_difference'].mean()
        std_diff = diff_df['temperature_difference'].std()
        min_diff = diff_df['temperature_difference'].min()
        max_diff = diff_df['temperature_difference'].max()
        
        stats_text = f'Mean: {mean_diff:.1f}°C, Std: {std_diff:.1f}°C\nRange: {min_diff:.1f}°C to {max_diff:.1f}°C'
        ax.text(0.02, 0.98, stats_text, transform=ax.transAxes,
                ha='left', va='top', fontsize=9,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='lightgreen', alpha=0.7))
        
        # Formatting
        ax.set_xlabel('Time')
        ax.set_ylabel('Temperature Difference (°C)')
        ax.set_title('Temperature Difference: Room (T3_Kek) - Intake (T1_BE)')
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        self.format_time_axis(ax)
    
    def plot_room_external_diff(self, ax, start_date: datetime, end_date: datetime):
        """Plot Room (T3_Kek) vs External (T2_Terasz) temperature difference."""
        # Get data for required devices
        room_df = self.db.get_device_data('T3_Kek', start_date, end_date)
        external_df = self.db.get_device_data('T2_Terasz', start_date, end_date)
        
        if room_df.empty or external_df.empty:
            ax.text(0.5, 0.5, 'Room vs External difference not available\nMissing T3_Kek or T2_Terasz data', 
                   transform=ax.transAxes, ha='center', va='center',
                   bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
            ax.set_title('Temperature Difference: Room - External')
            return
        
        # Synchronize the data by timestamp (simple approach)
        # Create a merged dataframe on timestamp with tolerance
        room_df = room_df[['timestamp', 'temperature']].rename(columns={'temperature': 'room_temp'})
        external_df = external_df[['timestamp', 'temperature']].rename(columns={'temperature': 'external_temp'})
        
        # Merge with tolerance (find nearest timestamps within 15 minutes)
        tolerance = pd.Timedelta(minutes=15)
        merged_data = []
        
        for _, room_row in room_df.iterrows():
            room_time = room_row['timestamp']
            room_temp = room_row['room_temp']
            
            # Find closest external temperature reading within tolerance
            time_diffs = abs(external_df['timestamp'] - room_time)
            closest_idx = time_diffs.idxmin()
            
            if time_diffs.loc[closest_idx] <= tolerance:
                external_temp = external_df.loc[closest_idx, 'external_temp']
                temp_diff = room_temp - external_temp
                
                merged_data.append({
                    'timestamp': room_time,
                    'room_temp': room_temp,
                    'external_temp': external_temp,
                    'temperature_difference': temp_diff
                })
        
        if not merged_data:
            ax.text(0.5, 0.5, 'No synchronized data available\nfor Room vs External comparison', 
                   transform=ax.transAxes, ha='center', va='center',
                   bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
            ax.set_title('Temperature Difference: Room - External')
            return
        
        # Convert to DataFrame for plotting
        diff_df = pd.DataFrame(merged_data)
        
        # Plot temperature difference
        ax.plot(diff_df['timestamp'], diff_df['temperature_difference'], 
               color='purple', alpha=0.8, linewidth=1.5, 
               label='Room - External (°C)')
        
        # Add horizontal line at zero
        ax.axhline(y=0, color='black', linestyle='--', alpha=0.5, linewidth=1)
        
        # Color areas above/below zero
        ax.fill_between(diff_df['timestamp'], diff_df['temperature_difference'], 0,
                      where=(diff_df['temperature_difference'] > 0), 
                      color='red', alpha=0.2, label='Room warmer')
        ax.fill_between(diff_df['timestamp'], diff_df['temperature_difference'], 0,
                      where=(diff_df['temperature_difference'] < 0), 
                      color='blue', alpha=0.2, label='External warmer')
        
        # Add statistics
        mean_diff = diff_df['temperature_difference'].mean()
        std_diff = diff_df['temperature_difference'].std()
        min_diff = diff_df['temperature_difference'].min()
        max_diff = diff_df['temperature_difference'].max()
        
        stats_text = f'Mean: {mean_diff:.1f}°C, Std: {std_diff:.1f}°C\nRange: {min_diff:.1f}°C to {max_diff:.1f}°C'
        ax.text(0.02, 0.98, stats_text, transform=ax.transAxes,
                ha='left', va='top', fontsize=9,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='lightblue', alpha=0.7))
        
        # Formatting
        ax.set_xlabel('Time')
        ax.set_ylabel('Temperature Difference (°C)')
        ax.set_title('Temperature Difference: Room (T3_Kek) - External (T2_Terasz)')
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        self.format_time_axis(ax)
    
    def get_data_type_label(self, data_type: str) -> str:
        """Get human-readable label for data type."""
        labels = {
            'temperature': 'Temperature (°C)',
            'humidity': 'Humidity (%)',
            'battery_mv': 'Battery (mV)'
        }
        return labels.get(data_type, data_type)
    
    def format_time_axis(self, ax):
        """Format time axis for better readability."""
        ax.tick_params(axis='x', rotation=45)
        
        # Auto-format dates based on range
        start_date = datetime.combine(self.start_date.get_date(), datetime.min.time())
        end_date = datetime.combine(self.end_date.get_date(), datetime.max.time())
        date_range = (end_date - start_date).days
        
        if date_range <= 7:
            # Show hours for short ranges
            import matplotlib.dates as mdates
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d %H:%M'))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=6))
        elif date_range <= 30:
            # Show days for medium ranges
            import matplotlib.dates as mdates
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
            ax.xaxis.set_major_locator(mdates.DayLocator(interval=3))
        else:
            # Show weeks for long ranges
            import matplotlib.dates as mdates
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
    
    def export_plot(self):
        """Export the current plot as an image."""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), 
                          ("SVG files", "*.svg"), ("All files", "*.*")],
                title="Export Plot"
            )
            
            if filename:
                self.fig.savefig(filename, dpi=300, bbox_inches='tight')
                self.update_status(f"Plot exported to {filename}")
                messagebox.showinfo("Export Successful", f"Plot saved as:\n{filename}")
                
        except Exception as e:
            logger.error(f"Error exporting plot: {e}")
            messagebox.showerror("Export Error", f"Failed to export plot:\n{str(e)}")
    
    def export_data(self):
        """Export the current data as CSV."""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), 
                          ("All files", "*.*")],
                title="Export Data"
            )
            
            if not filename:
                return
            
            # Get current selections
            selected_devices = self.get_selected_devices()
            start_date = datetime.combine(self.start_date.get_date(), datetime.min.time())
            end_date = datetime.combine(self.end_date.get_date(), datetime.max.time())
            
            # Combine all device data
            all_data = []
            for device in selected_devices:
                df = self.db.get_device_data(device, start_date, end_date)
                if not df.empty:
                    df['device'] = device
                    all_data.append(df)
            
            # Add STAT002 data if selected (calculate on-demand)
            if self.show_stat002_var.get():
                # Get data for T3_Kek and T1_BE
                room_df = self.db.get_device_data('T3_Kek', start_date, end_date)
                intake_df = self.db.get_device_data('T1_BE', start_date, end_date)
                
                if not room_df.empty and not intake_df.empty:
                    # Create synchronized temperature difference data
                    room_df = room_df[['timestamp', 'temperature']].rename(columns={'temperature': 'room_temp'})
                    intake_df = intake_df[['timestamp', 'temperature']].rename(columns={'temperature': 'intake_temp'})
                    
                    # Merge with tolerance (find nearest timestamps within 15 minutes)
                    tolerance = pd.Timedelta(minutes=15)
                    merged_data = []
                    
                    for _, room_row in room_df.iterrows():
                        room_time = room_row['timestamp']
                        room_temp = room_row['room_temp']
                        
                        # Find closest intake temperature reading within tolerance
                        time_diffs = abs(intake_df['timestamp'] - room_time)
                        if len(time_diffs) == 0:
                            continue
                            
                        closest_idx = time_diffs.idxmin()
                        
                        if time_diffs.loc[closest_idx] <= tolerance:
                            intake_temp = intake_df.loc[closest_idx, 'intake_temp']
                            temp_diff = room_temp - intake_temp
                            
                            merged_data.append({
                                'timestamp': room_time,
                                'device': 'STAT002_Room_Intake_Diff',
                                'temperature': temp_diff,
                                'humidity': None,
                                'battery_mv': None
                            })
                    
                    if merged_data:
                        diff_df = pd.DataFrame(merged_data)
                        all_data.append(diff_df)
            
            # Add Room vs External difference data if selected
            if self.show_room_external_diff_var.get():
                # Get data for T3_Kek and T2_Terasz
                room_df = self.db.get_device_data('T3_Kek', start_date, end_date)
                external_df = self.db.get_device_data('T2_Terasz', start_date, end_date)
                
                if not room_df.empty and not external_df.empty:
                    # Create synchronized temperature difference data
                    room_df = room_df[['timestamp', 'temperature']].rename(columns={'temperature': 'room_temp'})
                    external_df = external_df[['timestamp', 'temperature']].rename(columns={'temperature': 'external_temp'})
                    
                    # Merge with tolerance (find nearest timestamps within 15 minutes)
                    tolerance = pd.Timedelta(minutes=15)
                    merged_data = []
                    
                    for _, room_row in room_df.iterrows():
                        room_time = room_row['timestamp']
                        room_temp = room_row['room_temp']
                        
                        # Find closest external temperature reading within tolerance
                        time_diffs = abs(external_df['timestamp'] - room_time)
                        closest_idx = time_diffs.idxmin()
                        
                        if time_diffs.loc[closest_idx] <= tolerance:
                            external_temp = external_df.loc[closest_idx, 'external_temp']
                            temp_diff = room_temp - external_temp
                            
                            merged_data.append({
                                'timestamp': room_time,
                                'device': 'Room_External_Diff',
                                'temperature': temp_diff,
                                'humidity': None,
                                'battery_mv': None
                            })
                    
                    if merged_data:
                        diff_df = pd.DataFrame(merged_data)
                        all_data.append(diff_df)
            
            if not all_data:
                messagebox.showwarning("No Data", "No data to export with current selections.")
                return
            
            # Combine and save
            combined_df = pd.concat(all_data, ignore_index=True)
            combined_df = combined_df.sort_values(['device', 'timestamp'])
            
            if filename.endswith('.xlsx'):
                combined_df.to_excel(filename, index=False)
            else:
                combined_df.to_csv(filename, index=False)
            
            self.update_status(f"Data exported to {filename}")
            messagebox.showinfo("Export Successful", 
                              f"Data exported successfully:\n{filename}\n"
                              f"Records: {len(combined_df)}")
            
        except Exception as e:
            logger.error(f"Error exporting data: {e}")
            messagebox.showerror("Export Error", f"Failed to export data:\n{str(e)}")


def main():
    """Main application entry point."""
    try:
        # Check if database exists
        db_path = Path("data/temperature_database.json")
        if not db_path.exists():
            messagebox.showerror("Database Not Found", 
                               f"Temperature database not found at: {db_path}\n\n"
                               "Please run the data importer first:\n"
                               "python src/data_importer.py")
            return
        
        # Create and run GUI
        root = tk.Tk()
        app = TemperatureGUI(root)
        
        # Handle window closing
        def on_closing():
            root.quit()
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
        
    except Exception as e:
        logger.error(f"Application error: {e}")
        messagebox.showerror("Application Error", f"Failed to start application:\n{str(e)}")


if __name__ == "__main__":
    main()
