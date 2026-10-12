# Temperature Monitoring Project Overview

## 1. Project Overview & Context

This project tracks temperature and humidity data across various locations in a house using 8 Bluetooth-enabled data logging sensors.

- **Data frequency & capacity:** Sensors record data every 5 minutes. They store up to approximately 3 months of continuous historical data locally.
- **Operation history:** The system has been active for several years with minimal interruptions (primarily brief gaps caused by depleted batteries).

## 2. Sensor Placement & Locations

### Interior / living spaces
- "Kék szoba" (Blue Room): now used as a children's bedroom ("gyerekszoba").
- Living room ("Nappali")
- Office / study ("Dolgozó")

### Technical & utility spaces
- Shed / outbuilding ("Félszer")
- Cellar ("Pince")
- Pantry floor ("Kamra padló")
- Ventilation supply air outlet ("Szellőztető rendszer befújónyílása"): located specifically in the Blue Room.
- Radiator surface (heating zone 2): mounted on a radiator to track when Zone 2 heating is active based on temperature spikes.

### Outdoor & temporary locations
- Terrace ("Terasz"): measures ambient outdoor temperature.
- Temporary / historical locations: short-term datasets exist for other locations, such as inside a well ("kút").

## 3. Heating & Ventilation Systems Dynamics

### Heating system (Zones 1 & 2)

**Zone 1 (Z1 – underfloor heating)**
- Covers the central area of the house and the kitchen.
- *Behavior & override:* Due to airflow from the ventilation system pulling warm air toward central extraction points (e.g., bathrooms/kitchen), Zone 1 underfloor heating triggers infrequently on its own. To maintain optimal efficiency and comfort, a scheduled forced run (20-minute boost twice a day) is programmed via thermostat temperature setpoint adjustments.
- *Thermal inertia:* Concrete floor slabs have high heat capacity. Peak surface/room temperatures occur after active heating cycles conclude.

**Zone 2 (Z2 – radiator heating)**
- Covers living rooms/bedrooms.
- Monitored via the radiator-attached sensor to identify active heating cycles and gas consumption correlates.

### Ventilation system ("Szellőztető rendszer")
- **Sensor tag ("BE"):** sensor positioned at the supply vent inlet.
- **Airflow diagnostics:**
  - When active, supply air temperature trends closer to outdoor temperatures.
  - When switched off (e.g., during winter smoke/smog peaks or high summer heat), warm room air rises into the ductwork, causing the supply sensor to record rising temperatures. This serves as an indicator of whether ventilation was active or inactive.

## 4. Data Pipeline & Processing Requirements

### Data collection & storage
- Data is downloaded periodically via Bluetooth and consolidated in Google Drive.
- Files are archived together into ZIP archives.
- Raw data contains overlapping timestamps due to persistent sensor memory across downloads.
- Timestamps: recorded as standard Unix timestamps (5-minute log intervals across devices, with minor phase shifts/delays between individual units).

### Data preprocessing requirements
- **ZIP archive support:** processing scripts must ingest and parse ZIP archives directly.
- **Alignment & interpolation:** synchronize non-aligned 5-minute intervals across all 8 sensors to a unified timestamp grid using interpolation where appropriate.
- **Central database:** store consolidated readings incrementally into a single structured file format (typically JSON).
- **Version control:** repository storage (e.g., GitHub) for consolidated datasets and processing code.

## 5. Analytics & Key Insights

**Gas consumption correlation**
- Discrete gas meter readings are recorded periodically in tabular format.
- Cumulative degree-day differences (ΔT = T_inside − T_outside) correlate strongly with actual gas consumption.
- Heating cycle counts (Z1/Z2 activation count) show a near-linear relationship with total gas consumption.

## 6. Long-Term Goals & Claude / LLM Skill Integration

- **Incremental data ingestion:** easily append new sensor logs and manual gas meter readings into the main database.
- **Cloud & automation workflows:** create a custom AI Skill / Agent Workflow (integrating Python scripts running in cloud environments) capable of handling data ingestion, running diagnostic scripts, and generating automated thermal/energy analytics.
- **Base repository:** build upon existing scripts in the temperature-monitoring repository.
