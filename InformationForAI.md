# Information for AI - Temperature Monitoring System

This file contain- **Generated Files** (all in `output/`):
  - `ventilation_timeline.png`: Temperature readings, gradients, status over time
  - `ventilation_status_summary.png`: Pie/bar charts of ON/OFF/Uncertain distribution  
  - `ventilation_confidence_analysis.png`: Confidence score analysis
  - `ventilation_daily_heatmap.png`: Status by day×hour (🟢=ON, 🔴=OFF, 🟠=uncertain)
  - `ventilation_confidence_heatmap.png`: Confidence levels by day×hour (green-yellow-red scale)
  - `ventilation_combined_heatmap.png`: Side-by-side status and confidence comparison
  - `ventilation_temperature_difference_heatmap.png`: Room-Intake temp difference (thermal colormap)
  - `ventilation_comprehensive_heatmap.png`: 4-panel comprehensive analysis viewtial information for AI assistants to quickly understand the Temperature Monitoring System codebase and its current state.

## 🏗️ Project Overview

**Purpose**: Python console application for processing IoT temperature/humidity sensor data from multiple devices in a house.

**Key Workflow**: ZIP files with CSV data → JSON database → Statistics & Visualizations → Reports

**Current Status**: Fully functional with core features implemented (July 28, 2025)

## 📁 Critical File Locations

### Core Application Files
- `src/main.py` - Main CLI application entry point
- `src/data_importer.py` - ZIP/CSV processing, database updates (IN001)
- `src/temperature_processor.py` - Core data processing and validation
- `src/temperature_statistics.py` - Statistical analysis (STAT001) 
- `src/simple_visualizer.py` - Clean visualization generation (VIZ001)
- `src/excel_exporter.py` - Excel report generation (VIZ002)

### Key Data Files
- `data/temperature_database.json` - Central JSON database (~57MB, 214k+ records)
- `data/TempLogs*.zip` - Input ZIP files with CSV data
- `output/` - Generated reports, charts, and analysis files

### Documentation
- `specification.md` - Comprehensive project specification (UPDATED)
- `requirements.txt` - Python dependencies
- `pyproject.toml` - Project configuration

## 🔧 Implementation Status

### ✅ IMPLEMENTED FEATURES

#### (IN001) Data Import - `data_importer.py`
- **Timestamp Filtering**: Records before 2020-01-01 are automatically filtered out
- **Duplicate Detection**: MD5 hash-based deduplication across ZIP files
- **Batch Processing**: Handles multiple `TempLogs*.zip` files
- **Database Structure**: JSON with metadata, devices, records, hashes
- **Usage**: `python src/data_importer.py`

#### (STAT001) Active Time Intervals - `temperature_statistics.py`
- **Gap Detection**: Identifies data gaps > 35 minutes
- **Interval Splitting**: Creates separate intervals for discontinuous data
- **Completeness Analysis**: Expected vs actual record counts
- **Text Reports**: Human-readable interval analysis
- **Usage**: `python src/temperature_statistics.py`

#### (STAT002) Ventilation Analysis - `temperature_statistics.py` (SIMPLIFIED)
- **Temperature Difference Analysis**: Uses T1_BE (intake), T3_Kek (room), T2_Terasz (external)
- **Focus**: Room-Intake temperature difference as primary ventilation indicator
- **Results**: 27,459 data points with statistical analysis (Mean: 1.98°C, Range: 17.23°C)
- **Synchronized Data**: 15-minute tolerance for sensor timing differences
- **Output**: Simplified JSON results focusing on physical temperature relationships
- **Usage**: Automatically runs after STAT001 in `python src/temperature_statistics.py`

#### (VIZ001) Simple Visualizations - `simple_visualizer.py`
- **Quick Overview**: Multi-panel dashboard 
- **Device Summary**: Statistical bar charts
- **Temperature Comparison**: Multi-device time-series
- **Individual Timelines**: 3-panel device charts (temp/humidity/battery)
- **Usage**: `python src/simple_visualizer.py`

#### (VIZ002) Excel Export - `excel_exporter.py`
- **Multi-sheet Workbooks**: Raw data + hourly aggregation
- **Embedded Charts**: Temperature, humidity, battery visualizations
- **Integrated**: Called automatically from main.py

#### (VIZ003) Temperature Difference Heatmap - `visualize_ventilation.py` (SIMPLIFIED)
- **Single Visualization**: Creates only the temperature difference heatmap
- **Physical Focus**: Shows Room (T3_Kek) - Intake (T1_BE) temperature difference over time
- **Thermal Colormap**: RdBu_r colormap for intuitive red/blue temperature visualization  
- **Day×Hour Matrix**: Rows=dates, Columns=hours, Colors=temperature difference
- **Generated File**: `output/ventilation_temperature_difference_heatmap.png`
- **Dependencies**: Requires `temperature_statistics.py` to be run first (creates JSON data)
- **Usage**: `python visualize_ventilation.py` (simplified, single output)

#### (VIZ004) Interactive GUI Application - `temperature_gui.py` (NEW)
- **Modern Interface**: tkinter-based GUI with matplotlib integration
- **Device Selection**: Multi-select checkboxes for all 15 available sensors
- **Date Range Control**: Calendar pickers with quick presets (7/30 days)
- **Multiple Data Types**: Temperature, humidity, battery, or all in subplots
- **STAT002 Integration**: Temperature difference visualization with color coding
- **Export Features**: High-quality plot export (PNG/PDF/SVG) and data export (CSV/Excel)
- **Interactive Navigation**: Zoom, pan, reset with matplotlib toolbar
- **Real-time Updates**: Instant plot refresh on selection changes
- **Usage**: `python temperature_gui.py` or via VS Code task "Run Temperature GUI"

#### (MAIN001) CLI Interface - `main.py`
- **Single ZIP Processing**: `python src/main.py data/sample_data.zip`
- **Automatic Reports**: Generates visualizations and Excel exports
- **Console Summary**: Key statistics output

## 🗃️ Database Structure

**File**: `data/temperature_database.json` (Large file - 57MB)

```json
{
  "metadata": {
    "created": "timestamp",
    "last_updated": "timestamp", 
    "version": "1.0.0",
    "total_records": 214161
  },
  "devices": {
    "T1_BE": {
      "device_name": "T1_BE",
      "first_seen": "timestamp",
      "records": [
        {
          "timestamp": "2024-10-05T19:13:58",
          "temperature": 20.5,
          "humidity": 65.2,
          "battery_mv": 2943,
          "hash": "md5_hash_string"
        }
      ],
      "record_hashes": ["hash1", "hash2"],
      "total_records": 29188,
      "last_updated": "timestamp"
    }
  },
  "import_history": []
}
```

## 📊 Current Data Stats (as of July 28, 2025)

- **Total Devices**: 15 (T1_BE, T3_Kek, T8_Z1, T2_Terasz, T6_Z2, T7_Dolgozo, T4_Nappali, T5_Lampa, T5_KamraLent, T7_Pince, ATC_EF_T7, ATC_F1_T3, T7_Kut, ATC_E9_T4, T7_AUX)
- **Total Records**: 639,585 records  
- **Date Range**: October 2024 to July 2025 (ongoing)
- **Data Quality**: 99.9% completeness on most devices
- **Database Size**: ~57MB JSON file
- **STAT002 Analysis**: 27,459 synchronized data points for ventilation analysis

## 📝 CSV Data Format

**Structure**:
```csv
T3_Kek                                    # Row 1: Device name
Time-Data=(A2/86400)+25569;Temp;Humi;Vbat # Row 2: Header
1791;13.72;86.58;2943                     # Row 3+: Unix timestamp;temp;humidity;battery_mv
```

**Important**: 
- Timestamps are Unix seconds (integer)
- Battery values are millivolts (e.g., 2943 = 2.943V)
- Pre-2020 timestamps are filtered out during import

## 🛠️ VS Code Tasks Available

```json
{
  "Run Temperature Monitoring App": "python src/main.py data/sample_data.zip",
  "Run Unit Tests": "pytest tests/ -v", 
  "Run Tests with Coverage": "pytest tests/ --cov=src --cov-report=html",
  "Install Dependencies": "pip install -r requirements.txt",
  "Create Sample ZIP": "python -c '...'"  // Creates test data
}
```

## ⚠️ Important Implementation Notes

### STAT002 Simplification (July 28, 2025)
- **Previous**: Complex multi-indicator system with ON/OFF/UNCERTAIN classification (43.7% uncertainty)
- **Current**: Simplified temperature difference analysis focusing on Room-Intake relationship
- **Benefit**: Clear physical interpretation without algorithmic complexity
- **Output**: Single temperature difference heatmap with statistical analysis

### Timestamp Filtering (Critical)
- **MUST filter pre-2020 records**: Implemented in `data_importer.py` line ~98
- **Code location**: `_process_device_data()` method
- **Filter condition**: `if record['timestamp'] < datetime(2020, 1, 1)`

### Database Limitations
- **Large file warning**: JSON database can be 50MB+ (VSCode won't sync files >50MB)
- **Memory usage**: Loading full database requires significant RAM
- **Performance**: Streaming approach used in processors

### Visualization Issues (Minor)
- **Emoji warnings**: Unicode emoji in plots may show font warnings on Windows
- **Not critical**: Charts still generate correctly
- **Location**: `simple_visualizer.py` matplotlib plots with emoji titles

## 🧪 Testing Framework

- **Framework**: pytest with coverage reporting
- **Test files**: `tests/test_*.py`
- **Coverage**: HTML reports in `output/coverage_html/`
- **Status**: Comprehensive unit tests implemented

## 🔄 Typical Development Workflow

1. **New data import**: Run `python src/data_importer.py`
2. **Generate statistics**: Run `python src/temperature_statistics.py` (includes STAT001 + simplified STAT002)
3. **Interactive visualization**: Run `python temperature_gui.py` (GUI application)
4. **Create temperature difference heatmap**: Run `python visualize_ventilation.py`
5. **Create other visualizations**: Run `python src/simple_visualizer.py`
6. **Process single ZIP**: Run `python src/main.py data/file.zip`
7. **Run tests**: Use VS Code task or `pytest tests/ -v`

## 🐛 Common Issues & Solutions

### "Database not found" Error
- **Cause**: Missing `data/temperature_database.json`
- **Solution**: Run `python src/data_importer.py` first

### "No ZIP files found" Warning  
- **Cause**: No `TempLogs*.zip` files in `data/` folder
- **Solution**: Ensure ZIP files follow naming convention

### Font/Emoji Warnings in Plots
- **Cause**: Windows font doesn't support Unicode emoji
- **Impact**: Cosmetic only, plots still generate
- **Solution**: Ignore warnings or remove emoji from titles

## 📈 Performance Benchmarks

- **Import Speed**: ~3 seconds for full ZIP processing
- **Visualization Generation**: ~5 seconds for complete suite
- **Database Query**: Nearly instant for statistics
- **Memory Usage**: Efficient streaming for large datasets

## 🔮 Future Development Areas

### Not Yet Implemented (from specification)
- `(STAT002-005)` Additional statistical analyses
- `(VIZ003-005)` Advanced visualization features
- Integration with external data sources (gas meter, weather)
- Web dashboard interface

### Architecture Ready For
- New statistics modules (extend `temperature_statistics.py`)
- Additional visualization types (extend visualizer modules)
- New export formats (alongside Excel exporter)
- External data integration

## 🎯 Quick Debugging Commands

```bash
# Check database status
python -c "import json; db=json.load(open('data/temperature_database.json')); print(f'Devices: {len(db[\"devices\"])}, Records: {sum(len(d[\"records\"]) for d in db[\"devices\"].values())}')"

# List available ZIP files
dir data\TempLogs*.zip

# Check output files
dir output

# Run specific visualization
python src/simple_visualizer.py

# Generate specific report
python src/temperature_statistics.py
```

## 📋 Code Quality Notes

- **Type Hints**: All functions have comprehensive type annotations
- **Error Handling**: Robust exception handling with logging
- **Documentation**: Docstrings for all major functions
- **Modularity**: Clean separation of concerns
- **Testability**: Functions designed for unit testing

---

**Last Updated**: July 28, 2025 (STAT002 Simplified + GUI Added)  
**Database Version**: 639,585 records from 15 devices  
**Implementation Status**: Core features complete with simplified STAT002 and interactive GUI, ready for extensions  
**Key Changes**: 
- STAT002 simplified to focus on temperature difference analysis only
- New interactive GUI application for data visualization and exploration
