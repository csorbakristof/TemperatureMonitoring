# Temperature Monitoring GUI Application

Interactive GUI application for visualizing temperature sensor data with advanced filtering and export capabilities.

## Features

### 📊 Data Visualization
- **Interactive Time Series Plots**: High-quality matplotlib visualizations
- **Multiple Data Types**: Temperature, humidity, battery voltage
- **Device Selection**: Multi-select checkboxes for up to 15 sensors
- **Date Range Filtering**: Calendar-based date pickers with quick presets
- **Derived Statistics**: STAT002 temperature difference analysis

### 🎛️ User Interface
- **Modern GUI**: Clean tkinter interface with organized control panels
- **Responsive Layout**: Adjustable plot area with navigation toolbar
- **Real-time Updates**: Instant plot refresh on selection changes
- **Status Feedback**: Real-time status updates and error handling

### 📁 Export Capabilities
- **Plot Export**: Save as PNG, PDF, or SVG with high DPI
- **Data Export**: Export filtered data as CSV or Excel
- **Multiple Formats**: Support for various output formats

## Installation & Setup

### Prerequisites
```bash
# Install required packages
pip install -r requirements.txt

# Ensure database exists
python src/data_importer.py  # If needed
python src/temperature_statistics.py  # For STAT002 data
```

### Launch Application
```bash
# Method 1: Direct launch
python temperature_gui.py

# Method 2: Using launcher
python launch_gui.py

# Method 3: VS Code Task
# Press Ctrl+Shift+P → "Tasks: Run Task" → "Run Temperature GUI"
```

## Usage Guide

### 1. Device Selection
- **Select Devices**: Use checkboxes to choose sensors to visualize
- **Pre-selected**: Key devices (T1_BE, T3_Kek, T2_Terasz) are selected by default
- **Bulk Actions**: "Select All" / "Deselect All" buttons for convenience
- **Scrollable List**: Handles 15+ devices with scrollable interface

### 2. Data Type Options
- **Temperature (°C)**: Primary sensor readings
- **Humidity (%)**: Relative humidity measurements  
- **Battery (mV)**: Sensor battery voltage levels
- **All in subplots**: Display all three types in stacked subplots

### 3. Date Range Selection
- **Calendar Pickers**: Start and end date selection
- **Quick Presets**: "Last 7 days" and "Last 30 days" buttons
- **Auto-formatting**: X-axis automatically adjusts to date range

### 4. Advanced Features
- **STAT002 Integration**: Real-time temperature difference analysis
  - Room (T3_Kek) - Intake (T1_BE) difference
  - Color-coded areas (red=room warmer, blue=intake warmer)
  - Calculated on-demand with timestamp synchronization (15-minute tolerance)
  - Includes statistical summary (mean, std, range)
- **Room vs External Difference**: Real-time temperature difference calculation
  - Room (T3_Kek) - External (T2_Terasz) difference
  - Calculated on-demand with timestamp synchronization (15-minute tolerance)
  - Includes statistical summary (mean, std, range)

### 5. Export Functions
- **Export Plot**: Save current visualization as image
- **Export Data**: Save filtered dataset as CSV/Excel
- **High Quality**: 300 DPI export for publication-ready plots

## Technical Details

### Database Integration
- **JSON Database**: Reads from `data/temperature_database.json`
- **Efficient Loading**: Streaming approach for large datasets
- **Date Filtering**: Server-side filtering for performance

### Plot Features
- **Interactive Navigation**: Zoom, pan, and reset capabilities
- **Auto-scaling**: Intelligent axis scaling based on data range
- **Multi-device Colors**: Automatic color assignment for devices
- **Grid & Labels**: Clean formatting with proper axis labels

### Error Handling
- **Graceful Degradation**: Handles missing data gracefully  
- **User Feedback**: Clear error messages and status updates
- **Data Validation**: Checks for database availability and data integrity

## Available Devices (Example)

The GUI automatically detects all available devices from your database:
1. **T1_BE** - Intake sensor
2. **T3_Kek** - Room sensor  
3. **T2_Terasz** - External sensor
4. **T8_Z1** - Zone 1 sensor
5. **T6_Z2** - Zone 2 sensor
6. **T7_Dolgozo** - Worker area sensor
7. **T4_Nappali** - Living room sensor
8. **T5_KamraLent** - Basement storage sensor
9. **T7_Pince** - Cellar sensor
10. **And more...** (dynamically loaded)

## Data Sources

### Raw Sensor Data
- **Database**: `data/temperature_database.json`
- **Format**: Timestamp, temperature, humidity, battery
- **Filtering**: Pre-2020 data automatically excluded

### Derived Statistics  
- **STAT002**: `output/STAT002_ventilation_analysis.json`
- **Analysis**: Temperature difference calculations
- **Dependencies**: Requires `temperature_statistics.py` execution

## Troubleshooting

### Common Issues

**"Database Not Found" Error**
```bash
# Solution: Import data first
python src/data_importer.py
```

**"No STAT002 Data" Message**
```bash
# Solution: Generate statistics
python src/temperature_statistics.py
```

**GUI Won't Start**
```bash
# Check dependencies
pip install -r requirements.txt

# Verify tkinter (usually pre-installed)
python -c "import tkinter; print('OK')"
```

### Performance Tips
- **Large Date Ranges**: Use smaller ranges for better responsiveness
- **Device Selection**: Limit to 5-10 devices for optimal performance
- **Export**: Use PNG for smaller files, PDF for vector graphics

## Development Notes

### Code Structure
- **`TemperatureDatabase`**: Database interface and data loading
- **`StatisticsLoader`**: STAT002 and derived data loading
- **`TemperatureGUI`**: Main GUI class with all interface logic

### Extension Points
- **New Data Types**: Add to `data_types` list and `plot_single_data_type()`
- **Additional Statistics**: Extend `StatisticsLoader` for new analyses
- **Export Formats**: Add to `export_plot()` and `export_data()` methods

### Dependencies
- **tkinter**: Standard GUI framework (built-in)
- **tkcalendar**: Enhanced date picker widgets
- **matplotlib**: High-quality plotting with tkinter backend
- **pandas**: Data manipulation and filtering
- **numpy**: Numerical operations and array handling

---

**GUI Version**: 1.0.0  
**Compatible with**: Temperature Monitoring System v1.0+  
**Platform**: Windows, macOS, Linux (Python 3.8+)
