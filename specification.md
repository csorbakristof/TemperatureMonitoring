# Temperature Monitoring System Specification

## Project Scope

This is a Python console application for processing temperature and humidity monitoring data from multiple small IoT devices placed throughout a house. The system provides comprehensive data analysis, visualization, and reporting capabilities to gain insights into thermal, heating, cooling, and ventilation properties.

### Key Features:
- Import and merge data from multiple ZIP archives containing CSV files
- Deduplicate records across overlapping datasets
- Generate statistical analyses and visualizations
- Export data to Excel with charts and evaluations
- Comprehensive unit testing framework
- Modular architecture for easy extension

## Using the Software

### Data Collection Workflow:
1. **Device Data**: Temperature/humidity sensors record data every ~5 minutes
2. **Download Cycle**: Data is downloaded every 1-2 months (devices store ~3 months of data)
3. **Data Storage**: Devices overwrite oldest entries automatically
4. **File Management**: CSV files are uploaded to Google Drive and downloaded as ZIP packages
5. **Processing**: Application processes multiple ZIP files, handling duplicates and file name conflicts

### Data Processing:
- The system reads all `TempLogs*.zip` files from the `data/` folder
- CSV files may be nested in subfolders within ZIP archives
- Duplicate detection prevents redundant data from overlapping downloads
- All data is merged into a central JSON database (`data/temperature_database.json`)
- Analysis and visualization tools work directly from the JSON database

### Output Capabilities:
- Interactive statistical reports
- Time-series visualizations with customizable date ranges
- Multi-device comparison charts
- Excel exports with embedded charts
- Gap analysis and data completeness reports

## Software Architecture

### Core Components:
- **Data Importer** (`data_importer.py`): Handles ZIP file processing, CSV parsing, and database updates
- **Temperature Processor** (`temperature_processor.py`): Core data processing and timestamp conversion
- **Statistics Module** (`temperature_statistics.py`): Statistical analysis and reporting functions
- **Simple Visualizer** (`simple_visualizer.py`): Clean, focused visualization generation
- **Excel Exporter** (`excel_exporter.py`): Excel report generation with charts
- **GUI Application** (`temperature_gui.py`): Interactive GUI for data visualization and exploration
- **Main Application** (`main.py`): Command-line interface and workflow orchestration

### Database Structure:
The central JSON database (`data/temperature_database.json`) contains:
```json
{
  "metadata": {
    "created": "timestamp",
    "last_updated": "timestamp", 
    "version": "1.0.0",
    "total_records": 214161
  },
  "devices": {
    "device_name": {
      "device_name": "string",
      "first_seen": "timestamp",
      "records": [
        {
          "timestamp": "ISO string",
          "temperature": float,
          "humidity": float,
          "battery_mv": int,
          "hash": "md5_hash"
        }
      ],
      "record_hashes": ["hash1", "hash2", ...],
      "total_records": int,
      "last_updated": "timestamp"
    }
  },
  "import_history": [...]
}
```

### Extensibility:
- **Modular Statistics**: New statistics can be added by implementing functions in `temperature_statistics.py`
- **Visualization Plugins**: New chart types can be added to visualizer modules
- **Export Formats**: Additional export formats can be implemented alongside Excel exporter
- **Data Sources**: System can be extended to import additional data sources (e.g., gas meter readings)

### Testing Framework:
- Comprehensive unit tests using pytest
- Test coverage reports generated in `output/coverage_html/`  
- Modular testing allows individual component validation
- Support for test-driven development methodology

## Development Methodology

### Test-Driven Development:
- **Unit Testing**: Comprehensive test suite using pytest
- **Coverage Analysis**: Test coverage tracking with HTML reports
- **Continuous Integration**: Tests can be run via VS Code tasks
- **Modular Testing**: Each component has dedicated test files

### Code Quality Standards:
- **Type Hints**: All functions include comprehensive type hints
- **PEP 8 Compliance**: Code follows Python style guidelines  
- **Docstrings**: Comprehensive documentation for all functions
- **Error Handling**: Robust error handling with informative logging

### Available VS Code Tasks:
- `Run Temperature Monitoring App`: Execute main application
- `Run Unit Tests`: Execute test suite with verbose output
- `Run Tests with Coverage`: Generate coverage reports
- `Install Dependencies`: Install required packages
- `Create Sample ZIP`: Generate test data

## Implemented Functions

This section documents all implemented functions with their identifiers for easy reference.

### Data Import Functions

#### (IN001) Importing new data
**Status**: ✅ Implemented  
**Module**: `data_importer.py`

Processes ZIP files containing CSV data and updates the central JSON database:
- **ZIP Processing**: Extracts CSV files from `TempLogs*.zip` archives
- **Duplicate Detection**: Uses MD5 hashing to prevent duplicate records
- **Timestamp Filtering**: Automatically skips records before January 1, 2020 (Unix timestamp 1577836800)
- **Batch Processing**: Handles multiple ZIP files in sequence
- **Statistics Tracking**: Reports new records, duplicates, and filtered records

**Usage**: `python src/data_importer.py`

**Output**: Updates `data/temperature_database.json` with import statistics

### Statistical Analysis Functions

#### (STAT001) Active time interval of devices
**Status**: ✅ Implemented  
**Module**: `temperature_statistics.py`

Analyzes data continuity and identifies active time intervals for each device:
- **Gap Detection**: Identifies data gaps larger than 35 minutes
- **Interval Splitting**: Creates separate intervals for discontinuous data
- **Completeness Analysis**: Calculates expected vs actual record counts
- **Text Report**: Generates human-readable interval report

**Output Format**:
```
Device: T1_BE
----------------------------------------
  Interval 1:
    Start:      2024-10-05 19:13:58
    End:        2024-11-03 15:29:13  
    Duration:   28.8 days (692.25 hours)
    Records:    8323 / 8308 expected (100.2% complete)
```

**Usage**: `python src/temperature_statistics.py`

#### (STAT002) Ventillation periods
**Status**: ✅ Implemented  
**Module**: `temperature_statistics.py`

Temperature gradient analysis to estimate ventilation ON/OFF status using statistical inference.

**Algorithm**: Analyzes relationships between three key temperature sensors:
- `T1_BE` (ventilation intake) 
- `T3_Kek` (room temperature)
- `T2_Terasz` (external temperature)

**Detection Method**: Uses multiple indicators to estimate ventilation status:
1. **Intake follows external**: T1_BE temperature close to T2_Terasz (external air being drawn in)
2. **Significant room gradient**: Notable temperature difference between T1_BE and T3_Kek  
3. **High intake variability**: T1_BE shows more temperature variation than external sensor
4. **Intake-room correlation**: Temperature changes in T1_BE correlate with T3_Kek changes

**Parameters**:
- Window size: 60 minutes (configurable)
- Gradient threshold: 0.5°C (configurable)
- Synchronization tolerance: 15 minutes between sensor readings

**Output**: 
- Detailed JSON report with timestamp-by-timestamp analysis
- Summary statistics showing ON/OFF/UNCERTAIN periods
- Confidence scores for each estimation
- Temperature gradients and correlation metrics

**Current Results** (from latest run):
- Total data points: 27,447 over 296 days
- Ventilation ON: 18.8% of time
- Ventilation OFF: 37.5% of time  
- Uncertain: 43.7% of time
- Average confidence: 0.59

**Usage**: `python src/temperature_statistics.py` (automatically runs STAT002 after STAT001)

**Visualization**: 
- `python visualize_ventilation.py` creates timeline and summary charts
- `python create_heatmap.py` creates daily heatmap visualizations

**Generated Visualization Files**:
- `ventilation_timeline.png`: Temperature readings, gradients, and status over time
- `ventilation_status_summary.png`: Pie chart and bar chart distribution
- `ventilation_confidence_analysis.png`: Confidence score analysis
- `ventilation_daily_heatmap.png`: Daily status heatmap (days × hours)
- `ventilation_confidence_heatmap.png`: Daily confidence heatmap (days × hours)
- `ventilation_combined_heatmap.png`: Combined status and confidence view
- `ventilation_temperature_difference_heatmap.png`: Room-Intake temperature difference heatmap
- `ventilation_comprehensive_heatmap.png`: 4-panel comprehensive analysis

### Visualization Functions

#### (VIZ001) Simple Device Visualizations
**Status**: ✅ Implemented  
**Module**: `simple_visualizer.py`

Creates clean, focused visualizations from the JSON database:
- **Quick Overview**: Multi-panel dashboard with key statistics
- **Device Summary**: Bar charts showing averages and record counts
- **Temperature Comparison**: Multi-device time-series comparison
- **Individual Timelines**: Detailed 3-panel charts (temperature, humidity, battery)

**Generated Files**:
- `quick_overview.png`: Summary dashboard
- `device_summary_chart.png`: Statistical comparison
- `temperature_comparison_full.png`: Multi-device comparison
- `{device_name}_timeline_full.png`: Individual device charts

**Usage**: `python src/simple_visualizer.py`

#### (VIZ002) Excel Export with Charts  
**Status**: ✅ Implemented
**Module**: `excel_exporter.py`

Exports device data to Excel workbooks with embedded charts:
- **Raw Data**: Complete dataset with timestamps
- **Hourly Aggregation**: Statistical summaries by hour
- **Embedded Charts**: Temperature, humidity, and battery charts
- **Multi-sheet Format**: Organized data presentation

**Usage**: Integrated into main application workflow

#### (VIZ003) Interactive GUI Application
**Status**: ✅ Implemented
**Module**: `temperature_gui.py`

Modern GUI application for interactive data visualization and exploration:
- **Device Selection**: Multi-select checkboxes for all available sensors
- **Date Range Filtering**: Calendar-based date pickers with quick presets
- **Multiple Data Types**: Temperature, humidity, battery voltage visualizations
- **STAT002 Integration**: Real-time temperature difference analysis (T3_Kek - T1_BE) calculated on-demand
- **Room vs External Difference**: Real-time temperature difference calculation (T3_Kek - T2_Terasz)
- **Export Capabilities**: High-quality plot and data export (PNG, PDF, SVG, CSV, Excel)
- **Interactive Navigation**: Zoom, pan, and reset functionality
- **Real-time Updates**: Instant plot refresh on selection changes

**Features**:
- Scrollable device selection for 15+ sensors
- Intelligent axis formatting based on date range
- Color-coded temperature difference analysis
- Professional-quality matplotlib integration
- Graceful error handling and status feedback

**Usage**: `python temperature_gui.py` or `python launch_gui.py`

### Main Application Functions

#### (MAIN001) Command-Line Interface
**Status**: ✅ Implemented  
**Module**: `main.py`

Primary application interface for processing ZIP files:
- **ZIP Processing**: Handles individual ZIP file analysis
- **Report Generation**: Automatic visualization and Excel export
- **Summary Statistics**: Console output with key metrics
- **Error Handling**: Comprehensive error reporting

**Usage**: `python src/main.py data/sample_data.zip`

## Data Format Specifications

### ZIP File Structure
The system processes ZIP files with the naming pattern `TempLogs*.zip` from the `data/` directory:
```
TempLogs20241030.zip
├── 20241030/
│   ├── data.csv
│   ├── data (1).csv  
│   ├── data (2).csv
│   └── ...
```

### CSV File Format
Each CSV file contains data for a single device with the following structure:

**Row 1**: Device name only
```
T3_Kek
```

**Row 2**: Header row (format indicator)
```
Time-Data=(A2/86400)+25569;Temp;Humi;Vbat
```

**Row 3+**: Data records with columns:
- **Timestamp**: Integer value (Unix timestamp in seconds)
- **Temperature**: Float (degrees Celsius)  
- **Humidity**: Float (percentage 0-100)
- **Battery**: Integer (voltage in millivolts, e.g., 2943 = 2.943V)

**Example**:
```csv
T3_Kek
Time-Data=(A2/86400)+25569;Temp;Humi;Vbat
1791;13.72;86.58;2943
3593;13.57;86.7;2957
5396;13.5;86.88;2960
7198;13.49;87.25;2960
9000;13.49;87.35;2961
10802;13.42;87.46;2962
12605;13.34;87.58;2962
14407;13.06;86.97;2963
16209;19.31;75.52;2964
18012;22.31;67.15;2975
19814;22.52;65.71;2981
21616;22.58;65.17;2984
```

### Timestamp Processing
- **Input Format**: Unix timestamps (seconds since 1970-01-01 UTC)
- **Validation**: Timestamps must be between 1970 and 2100 for acceptance
- **Filtering**: Records before 2020-01-01 are automatically rejected
- **Conversion**: Internal storage uses ISO 8601 format strings

### Device Naming Convention
Device names typically follow the pattern `T{number}_{location}`:
- `T1_BE`: Temperature sensor 1 in bedroom
- `T3_Kek`: Temperature sensor 3 in kitchen  
- `T8_Z1`: Temperature sensor 8 in zone 1
- `T2_Terasz`: Temperature sensor 2 on terrace
- `T5_KamraLent`: Temperature sensor 5 in basement storage

## System Requirements

### Dependencies
Core Python packages (see `requirements.txt`):
```
matplotlib>=3.5.0
seaborn>=0.11.0  
pandas>=1.5.0
openpyxl>=3.0.0
pytest>=7.0.0
pytest-cov>=4.0.0
```

### File Structure
```
TemperatureMonitoring/
├── data/                          # Input ZIP files and database
│   ├── TempLogs*.zip             # Input data files
│   ├── temperature_database.json # Central database
│   └── extracted/                # Temporary extraction folder
├── src/                          # Source code
│   ├── __init__.py
│   ├── main.py                   # Main application
│   ├── data_importer.py          # Data import (IN001)
│   ├── temperature_processor.py   # Core processing
│   ├── temperature_statistics.py # Statistics (STAT001)
│   ├── simple_visualizer.py      # Visualizations (VIZ001)
│   └── excel_exporter.py         # Excel export (VIZ002)
├── tests/                        # Unit tests
│   ├── test_*.py                 # Test files
│   └── __init__.py
├── output/                       # Generated reports and charts
│   ├── *.png                     # Visualization files
│   ├── *.xlsx                    # Excel reports
│   ├── *.txt                     # Text reports
│   └── coverage_html/            # Test coverage reports
├── requirements.txt              # Python dependencies
├── pyproject.toml               # Project configuration
├── specification.md             # This specification
└── README.md                    # User documentation
```

### Performance Characteristics
Based on current implementation:
- **Database Size**: ~57MB for 214,161 records from 9 devices
- **Processing Speed**: ~3 seconds for full ZIP import
- **Memory Usage**: Efficient streaming processing for large datasets
- **Visualization Generation**: ~5 seconds for complete chart suite

## Future Enhancements

### Planned Statistics (Not Yet Implemented)
- `(STAT002)` Temperature trend analysis
- `(STAT003)` Heating/cooling cycle detection  
- `(STAT004)` Energy consumption correlation
- `(STAT005)` Anomaly detection and alerts

### Planned Visualizations
- `(VIZ003)` Interactive web dashboard
- `(VIZ004)` Advanced statistical charts
- `(VIZ005)` Comparative analysis tools

### Integration Possibilities
- Gas meter data correlation
- Weather data integration
- Home automation system integration
- Mobile app companion
