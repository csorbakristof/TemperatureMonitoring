# Temperature Monitoring Data Processor

A Python console application for processing temperature sensor data from CSV files packed in ZIP archives. The application generates comprehensive visualizations and Excel reports for temperature monitoring analysis.

## Features

- **CSV Data Processing**: Parse CSV files with device names, timestamps, temperature, humidity, and battery data
- **ZIP File Support**: Extract and process multiple CSV files from ZIP archives
- **Data Visualization**: Generate timeline plots, multi-device comparisons, and statistical heatmaps
- **Excel Export**: Create detailed Excel reports with statistics and hourly averages
- **Comprehensive Testing**: Full unit test coverage with pytest
- **Logging**: Detailed logging for debugging and monitoring

## CSV File Format

The application expects CSV files with the following format:

```csv
T3_DeviceName
Time-Data=(A2/86400)+25569;Temp;Humi;Vbat
1791;13.72;86.58;2943
3593;13.57;86.7;2957
5396;13.5;86.88;2960
```

- **Line 1**: Device name (e.g., "T3_DeviceName")
- **Line 2**: Headers (Time-Data, Temp, Humi, Vbat)
- **Data Lines**: Timestamp (Excel format), Temperature (°C), Humidity (%), Battery (mV)

## Installation

### Quick Setup (Recommended)

1. **Clone or download** this project to your local machine

2. **Run the automated setup script**:
   ```powershell
   # Full setup with package verification
   .\setup.ps1
   
   # Quick setup without verification tests
   .\setup.ps1 -SkipTests
   
   # Show help and options
   .\setup.ps1 -Help
   ```

The setup script will automatically:
- Detect Python 3.8+ installation
- Create and activate a virtual environment
- Install all required dependencies
- Create necessary data and output directories
- Verify package installations (unless `-SkipTests` is used)

### Manual Installation

If you prefer manual setup or the automated script doesn't work:

1. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Create data directories** (if not already present):
   ```bash
   mkdir -p data output
   ```

## Usage

### Console Application

After running the setup script, you can start the application directly:

```bash
# Default: Visualizations + save to database
python src/main.py data/temperature_data.zip

# Add Excel reports (time-consuming)
python src/main.py data/temperature_data.zip --excel-reports

# Skip database saving (faster for one-time analysis)
python src/main.py data/temperature_data.zip --no-database

# Process without visualizations or Excel (database only)
python src/main.py data/temperature_data.zip --no-reports

# Full processing: visualizations + Excel + database
python src/main.py data/temperature_data.zip --excel-reports

# Fastest: no reports, no database
python src/main.py data/temperature_data.zip --no-reports --no-database

# Set logging level
python src/main.py data/temperature_data.zip --log-level DEBUG
```

**Note**: 
- **Database saving** is **enabled by default** - processed data is automatically saved to `data/temperature_database.json` for future analysis and to prevent reprocessing the same data
- **Excel reports** (`--excel-reports`) are **disabled by default** as they can be time-consuming for large datasets
- Use `--no-database` for faster one-time analysis when you don't need persistent storage

### GUI Application (if available)

To start the graphical user interface:

```bash
# Launch the GUI application
python launch_gui.py
```

The GUI provides:
- File selection dialog for ZIP files
- Visual progress indicators
- Interactive report generation options
- Real-time log display

### Simple Database Visualizer

For clean, presentation-ready visualizations from your existing JSON database:

```bash
# Generate overview charts from database (requires existing database)
python src/simple_visualizer.py
```

**Note**: This tool requires a JSON database created by `main.py` first. It generates:
- Quick overview with summary statistics
- Device comparison charts
- Clean individual device timelines
- Professional-quality visualizations optimized for presentations

### Heating Detection Analysis

Analyze temperature data to detect heating cycles in different zones:

```bash
# Detect heating cycles from temperature database
python detect_heating.py
```

**Note**: This tool requires a JSON database created by `main.py` first. It analyzes temperature patterns to identify when heating systems are active, using the following logic:
- **Heating starts**: Temperature rises 5°C above the daily minimum
- **Heating ends**: Temperature drops 1°C below the cycle maximum
- **Gap merging**: Cycles separated by less than 15 minutes are merged

**Outputs generated**:
- `heating_cycles.json`: JSON file with detected heating cycles
- `heating_cycle_per_day_Z1.png` / `heating_cycle_per_day_Z2.png`: Charts showing number of cycles per day
- `heating_duration_per_day_Z1.png` / `heating_duration_per_day_Z2.png`: Charts showing heating duration per day
- `heating_cycle_per_day_Z1.csv` / `heating_cycle_per_day_Z2.csv`: CSV data for cycles per day
- `heating_duration_per_day_Z1.csv` / `heating_duration_per_day_Z2.csv`: CSV data for duration per day
- `heating_analysis_summary.txt`: Summary report with statistics

### Calendar Image Generation

Create calendar heatmap visualizations for temperature and heating activity:

```bash
# Generate calendar heatmaps from database and heating cycles
python generate_calendars.py
```

**Note**: This tool requires both `temperature_database.json` (from `main.py`) and `heating_cycles.json` (from `detect_heating.py`). It creates calendar-style heatmaps where:
- **Rows**: Days
- **Columns**: Time of day (288 samples per day, 5-minute intervals)
- **Colors**: Temperature values (viridis colormap) or heating activity (red/white)

**Outputs generated**:
- `TempCal_Z1.png` / `TempCal_Z2.png`: Temperature calendars for Zone 1 and Zone 2
- `Heating_Z1.png` / `Heating_Z2.png`: Heating activity calendars (red=heating, white=off)
- `TempDiff.png`: Temperature difference calendar (T3_Kek - T2_Terasz)
- `Temp_Outside.png`: Outside temperature calendar (T2_Terasz)

### Heating Statistics Analysis

Analyze heating patterns and their relationship to temperature differences:

```bash
# Analyze heating statistics from database and heating cycles
python heating_statistics.py
```

**Note**: This tool requires both `temperature_database.json` (from `main.py`) and `heating_cycles.json` (from `detect_heating.py`). It analyzes the correlation between indoor/outdoor temperature differences and heating activity:
- Compares internal temperature (T3_Kek) vs external temperature (T2_Terasz)
- Correlates temperature differences with heating cycle frequency
- Includes days with zero heating cycles as valid data points

**Outputs generated**:
- `MeanDiff_And_HeatingCycleCount_Plot.png`: Dual-axis time series plot (temperature difference vs heating cycles)
- `MeanDiff_And_HeatingCycleCount_XY.png`: Scatter plot showing correlation
- `MeanDiff_And_HeatingCycleCount_Plot.csv`: CSV data export
- `MeanDiff_And_HeatingCycleCount_XY.csv`: CSV data export
- `heating_statistics.txt`: Summary statistics and correlation analysis

### Gas Meter Data Loader

Load gas meter readings from CSV files into the temperature database:

```bash
# Load gas meter data from CSV file
python loadGasmeterValuesIntoDatabase.py data/gasmeter_readings.csv

# Show help and usage information
python loadGasmeterValuesIntoDatabase.py --help
```

**CSV file format**:
- **Column 1**: Date (MM/DD/YYYY)
- **Column 2**: Time (HH:MM)
- **Column 3**: Gas meter value (float)

**Note**: This tool loads gas meter readings and adds them to `temperature_database.json` under a `gasmeter` section alongside `devices`. Features include:
- Automatic header detection and skipping
- Duplicate detection (skips records with same timestamp)
- Automatic sorting by timestamp
- Metadata tracking (unit: m³, last updated, total records)

**Database structure**:
The gas meter data is stored in the database as:
```json
{
  "gasmeter": {
    "records": [
      {"timestamp": "2024-11-01T07:38:00", "value": 5054.83},
      ...
    ],
    "metadata": {
      "description": "Gas meter readings",
      "unit": "m³",
      "last_updated": "...",
      "total_records": 33
    }
  }
}
```

### Available Tools Summary

The Temperature Monitoring system provides several specialized tools:

| Tool | Purpose | Input | Key Features |
|------|---------|-------|--------------|
| `main.py` | **Primary processor** | ZIP files | Data processing, database creation, visualizations |
| `simple_visualizer.py` | **Clean visualizations** | JSON database | Professional charts, overview summaries |
| `detect_heating.py` | **Heating analysis** | JSON database | Cycle detection, daily statistics, duration charts |
| `generate_calendars.py` | **Calendar heatmaps** | JSON database + heating cycles | Calendar visualizations, temperature/heating patterns |
| `heating_statistics.py` | **Statistical analysis** | JSON database + heating cycles | Correlation analysis, temperature difference trends |
| `loadGasmeterValuesIntoDatabase.py` | **Gas meter loader** | CSV files | Load gas meter readings into database |
| `data_importer.py` | **Batch processing** | Multiple ZIP files | Database building, import statistics |
| `setup.ps1` | **Environment setup** | None | Automated dependency installation |

**Typical workflow:**
1. Run `setup.ps1` (one-time setup)
2. Process data: `python src/main.py data/file.zip` (creates database)
3. Load gas meter data: `python loadGasmeterValuesIntoDatabase.py data/gasmeter.csv` (optional)
4. Generate clean charts: `python src/simple_visualizer.py` (uses database)
5. Detect heating cycles: `python detect_heating.py` (analyzes heating patterns)
6. Generate calendar heatmaps: `python generate_calendars.py` (creates calendar visualizations)
7. Analyze heating statistics: `python heating_statistics.py` (correlation analysis)
8. Temperature difference heatmap (room vs ventillation intake):
   - ./src/temperature_statistics.py
   - create_heatmap.py
9. Temperature interactive GUI: .\temperature_gui.py (or launch_gui.py)


## Program Outputs

The Temperature Monitoring application generates comprehensive reports and visualizations in the `output/` directory. Here's a complete summary of all outputs and their purposes:

### Individual Device Reports

**Timeline Visualization**
- **File**: `{device_name}_timeline.png`
- **Content**: Multi-panel timeline showing temperature, humidity, and battery voltage over time for a single device
- **Purpose**: Track individual device performance and identify patterns or anomalies

**Excel Data Report** *(Generated only with `--excel-reports` flag)*
- **File**: `{device_name}_data.xlsx`
- **Content**: 
  - Raw data with converted timestamps
  - Statistical summary (min, max, mean, std deviation)
  - Hourly averages for trend analysis
- **Purpose**: Detailed data analysis and further processing in Excel

### Multi-Device Comparison Reports

**Temperature Comparison**
- **File**: `multi_device_temperature_comparison.png`
- **Content**: Overlaid temperature timelines for all devices with different colors
- **Purpose**: Compare temperature readings across multiple sensors to identify environmental variations

**Humidity Comparison**
- **File**: `multi_device_humidity_comparison.png`
- **Content**: Overlaid humidity timelines for all devices
- **Purpose**: Analyze humidity patterns and correlations between different monitoring locations

**Battery Level Comparison**
- **File**: `multi_device_battery_mv_comparison.png`
- **Content**: Battery voltage timelines showing power consumption patterns
- **Purpose**: Monitor device health and predict maintenance needs

### Statistical Analysis

**Statistics Heatmap**
- **File**: `statistics_heatmap.png`
- **Content**: Color-coded matrix showing statistical measures (mean, std, min, max) for all devices
- **Purpose**: Quick visual comparison of device performance characteristics

**Combined Excel Report** *(Generated only with `--excel-reports` flag)*
- **File**: `all_devices_data.xlsx`
- **Content**: 
  - Consolidated data from all devices
  - Cross-device statistical comparisons
  - Summary statistics table
- **Purpose**: Comprehensive analysis and reporting for all monitored devices

### Log Files

**Application Log**
- **File**: `temperature_monitoring.log`
- **Content**: 
  - Processing steps and timestamps
  - Error messages and warnings
  - Data validation results
  - Performance metrics
- **Purpose**: Debugging, monitoring, and audit trail

**JSON Database** *(Generated by default)*
- **File**: `data/temperature_database.json`
- **Content**: 
  - All processed temperature data from devices
  - Import history and statistics
  - Metadata about processing sessions
  - Duplicate detection records
- **Purpose**: Persistent storage, historical analysis, and preventing data reprocessing

### Console Output

During execution, the application displays:
- **Processing Status**: Files being processed and progress indicators
- **Data Summary**: Device count, data points, and time range
- **Validation Results**: Invalid data lines and error counts
- **Report Generation**: Files created and their locations
- **Statistics Overview**: Quick summary of key metrics for each device

### Error Handling Outputs

When issues occur, the application generates:
- **Error Messages**: Clear descriptions of problems encountered
- **Warning Messages**: Non-fatal issues (e.g., malformed data lines)
- **Skipped Files**: List of files that couldn't be processed
- **Validation Reports**: Details about data quality issues

All outputs are designed to provide both immediate insights and detailed data for further analysis, making the application suitable for both quick monitoring and comprehensive research purposes.

### Python API

```python
from src.main import TemperatureMonitoringApp

app = TemperatureMonitoringApp()
device_data_list = app.process_zip_file("data/temperature_data.zip")
app.print_summary(device_data_list)
```

### Individual Components

```python
# Process CSV files
from src.temperature_processor import TemperatureDataProcessor
processor = TemperatureDataProcessor()
data = processor.process_zip_file("data/sensors.zip")

# Create visualizations
from src.visualizer import TemperatureVisualizer
visualizer = TemperatureVisualizer()
visualizer.create_temperature_timeline(data[0])

# Export to Excel
from src.excel_exporter import ExcelExporter
exporter = ExcelExporter()
exporter.export_device_data(data[0])
```

## Generated Reports

The application generates the following outputs in the `output/` directory:

### Visualizations
- `{device_name}_timeline.png`: Individual device timeline plots
- `multi_device_temperature_comparison.png`: Temperature comparison across devices
- `multi_device_humidity_comparison.png`: Humidity comparison across devices
- `multi_device_battery_mv_comparison.png`: Battery level comparison across devices
- `statistics_heatmap.png`: Statistical overview heatmap

### Excel Reports
- `{device_name}_data.xlsx`: Individual device data with statistics and hourly averages
- `all_devices_data.xlsx`: Combined report for all devices

## Testing

Run the comprehensive test suite:

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov=src --cov-report=html

# Run specific test file
pytest tests/test_temperature_processor.py

# Run with verbose output
pytest -v
```

## Project Structure

```
TemperatureMonitoring/
├── src/
│   ├── main.py                 # Main application entry point
│   ├── temperature_processor.py # Core CSV/ZIP processing
│   ├── visualizer.py           # Data visualization
│   └── excel_exporter.py       # Excel report generation
├── tests/
│   ├── test_temperature_processor.py
│   ├── test_visualizer.py
│   └── test_excel_exporter.py
├── data/                       # Input data directory
├── output/                     # Generated reports directory
├── requirements.txt            # Python dependencies
├── pyproject.toml             # Project configuration
├── specification.md           # Project specifications
└── README.md                  # This file
```

## Dependencies

- **pandas**: Data manipulation and analysis
- **matplotlib**: Basic plotting and visualization
- **seaborn**: Statistical data visualization
- **openpyxl**: Excel file creation and formatting
- **pytest**: Testing framework
- **pytest-cov**: Test coverage reporting
- **numpy**: Numerical computing support

## Development

### Adding New Features

1. Create new modules in the `src/` directory
2. Add corresponding tests in the `tests/` directory
3. Update requirements.txt if new dependencies are needed
4. Update this README with usage examples

### Code Style

- Follow PEP 8 standards
- Use type hints for all functions
- Include comprehensive docstrings
- Write testable, modular code

## Logging

The application logs to both console and file (`temperature_monitoring.log`):

- **INFO**: Normal operation messages
- **WARNING**: Non-fatal issues (e.g., skipped invalid data lines)
- **ERROR**: Serious problems that prevent processing
- **DEBUG**: Detailed debugging information

## Error Handling

The application handles common error scenarios:

- Invalid ZIP files
- Malformed CSV data
- Missing files
- Empty datasets
- File permission issues

## License

This project is provided as-is for temperature monitoring data analysis purposes.

## Support

For questions or issues, please refer to the specification.md file or check the log files for detailed error information.
