# Overview

detect_heating.py loads the JSON database and creates statistics about the status of the heating. There are two separate heating zones monitored by the devices T8_Z1 (zone 1) and T6_Z2 (zone 2).

## Heating Detection Algorithm

The heating detection uses a sophisticated daily minimum and cycle maximum approach:

**Heating Start Detection:**
- Heating is activated when the device measures a temperature ≥ +5°C above the minimum temperature of that day
- Each day's minimum temperature is calculated from all measurements for that day
- This heuristic accounts for natural daily temperature variation

**Heating End Detection:**  
- During a heating cycle, the system tracks the maximum temperature reached
- The heating cycle ends when the measured temperature drops ≥ 1°C below the current cycle's maximum temperature
- This ensures the system detects when heating actually stops, not just temporary fluctuations

**Gap Merging:**
- If the time between two heating cycles is shorter than 15 minutes, they are merged into a single cycle
- This handles brief interruptions that don't represent true heating stops

## Configuration Parameters

- `TEMP_RISE_ABOVE_MIN = 5.0°C` - Temperature rise above daily minimum to trigger heating start
- `TEMP_DROP_BELOW_MAX = -1.0°C` - Temperature drop below cycle maximum to trigger heating end  
- `MIN_GAP_MINUTES = 15` - Minimum gap between cycles (shorter gaps are merged)

## Data Processing Rules

- Process every data point for the given devices in the database
- Missing data periods are treated as "inactive heating"
- All timestamps are in the same timezone (no timezone conversion needed)
- Weekdays and weekend days are not separated in the analysis
- If a heating cycle spans midnight, it belongs to the day it started (first day)
- Daily minimums are calculated per calendar date for threshold determination

## Output Files

### JSON Data Output
**heating_cycles.json** - Contains detailed heating cycle data for both zones:
```json
{
  "T8_Z1": [
    {
      "start": "2024-10-01T08:30:00", 
      "end": "2024-10-01T10:15:00", 
      "maxtemp": "34.5", 
      "durationMinutes": "105"
    },
    {
      "start": "2024-10-01T18:00:00", 
      "end": "2024-10-01T22:30:00", 
      "maxtemp": "32.2", 
      "durationMinutes": "270"
    }
  ],
  "T6_Z2": [...]
}
```

### Chart Outputs
**Daily Cycle Charts:**
- `heating_cycle_per_day_Z1.png` - Zone 1 cycles per day (dots only, no lines)
- `heating_cycle_per_day_Z2.png` - Zone 2 cycles per day (dots only, no lines)
- X-axis: Date, Y-axis: Number of heating cycles

**Daily Duration Charts:**  
- `heating_duration_per_day_Z1.png` - Zone 1 heating hours per day (dots only, no lines)
- `heating_duration_per_day_Z2.png` - Zone 2 heating hours per day (dots only, no lines)  
- X-axis: Date, Y-axis: Total heating duration in hours

### CSV Data Exports
For every PNG chart, a corresponding CSV file contains the raw data:
- `heating_cycle_per_day_Z1.csv` - Raw data for Z1 daily cycles
- `heating_cycle_per_day_Z2.csv` - Raw data for Z2 daily cycles  
- `heating_duration_per_day_Z1.csv` - Raw data for Z1 daily durations
- `heating_duration_per_day_Z2.csv` - Raw data for Z2 daily durations

### Summary Report
**heating_analysis_summary.txt** - Statistical summary including:
- Total cycles and heating hours per zone
- Average cycle duration and maximum temperature
- Daily statistics (max/average cycles and duration per day)
- Date range coverage
