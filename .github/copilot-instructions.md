<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Temperature Monitoring Application

This is a Python console application for processing temperature monitoring data from CSV files packed in ZIP archives.

## Key Features:
- Parse CSV files with device names, timestamps, temperature, humidity, and battery data
- Generate visualizations and Excel reports
- Comprehensive unit testing with pytest

## CSV Format:
- First row: Device name only
- Second row: Headers (Time-Data, Temp, Humi, Vbat)
- Data rows: Timestamp (Excel format), Temperature (°C), Humidity (%), Battery (mV)

## Code Style:
- Use type hints for all functions
- Follow PEP 8 standards
- Include comprehensive docstrings
- Write testable, modular code
