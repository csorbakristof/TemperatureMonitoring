#!/usr/bin/env python3
"""
Gas Meter Data Loader

Loads gas meter readings from a CSV file and adds them to the temperature database.
CSV format: date (MM/DD/YYYY), time (HH:MM), gasmeter value (float)
"""

import json
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Dict
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# Configuration
DATABASE_PATH = "data/temperature_database.json"


class GasmeterDataLoader:
    """Loads gas meter data from CSV and updates the temperature database."""
    
    def __init__(self, database_path: str = DATABASE_PATH):
        self.database_path = Path(database_path)
        self.database = self._load_database()
    
    def _load_database(self) -> Dict:
        """Load the JSON database."""
        if not self.database_path.exists():
            logger.error(f"Database not found: {self.database_path}")
            logger.error("Please run main.py first to create the temperature database.")
            sys.exit(1)
        
        try:
            with open(self.database_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse database JSON: {e}")
            sys.exit(1)
    
    def _save_database(self):
        """Save the updated database."""
        try:
            with open(self.database_path, 'w', encoding='utf-8') as f:
                json.dump(self.database, f, indent=2, ensure_ascii=False)
            logger.info(f"Database saved successfully: {self.database_path}")
        except Exception as e:
            logger.error(f"Failed to save database: {e}")
            sys.exit(1)
    
    def _parse_datetime(self, date_str: str, time_str: str) -> str:
        """Parse date and time strings into ISO format timestamp."""
        try:
            # Parse MM/DD/YYYY HH:MM
            datetime_str = f"{date_str} {time_str}"
            dt = datetime.strptime(datetime_str, "%m/%d/%Y %H:%M")
            return dt.isoformat()
        except ValueError as e:
            logger.error(f"Failed to parse date/time '{date_str} {time_str}': {e}")
            return None
    
    def load_csv_file(self, csv_path: str) -> List[Dict]:
        """Load gas meter data from CSV file."""
        csv_file = Path(csv_path)
        
        if not csv_file.exists():
            logger.error(f"CSV file not found: {csv_path}")
            sys.exit(1)
        
        records = []
        skipped_lines = 0
        
        try:
            with open(csv_file, 'r', encoding='utf-8') as f:
                # Try to detect if there's a header row
                first_line = f.readline().strip()
                f.seek(0)
                
                # Check if first line looks like a header
                has_header = any(keyword in first_line.lower() 
                               for keyword in ['date', 'time', 'meter', 'value', 'gas'])
                
                reader = csv.reader(f)
                
                if has_header:
                    next(reader)  # Skip header
                    logger.info("Detected header row, skipping it")
                
                for line_num, row in enumerate(reader, start=2 if has_header else 1):
                    if len(row) < 3:
                        logger.warning(f"Line {line_num}: Not enough columns (expected 3, got {len(row)})")
                        skipped_lines += 1
                        continue
                    
                    date_str = row[0].strip()
                    time_str = row[1].strip()
                    value_str = row[2].strip()
                    
                    # Skip empty lines
                    if not date_str or not time_str or not value_str:
                        skipped_lines += 1
                        continue
                    
                    # Parse timestamp
                    timestamp = self._parse_datetime(date_str, time_str)
                    if timestamp is None:
                        skipped_lines += 1
                        continue
                    
                    # Parse gas meter value
                    try:
                        value = float(value_str)
                    except ValueError:
                        logger.warning(f"Line {line_num}: Invalid gas meter value '{value_str}'")
                        skipped_lines += 1
                        continue
                    
                    records.append({
                        'timestamp': timestamp,
                        'value': value
                    })
        
        except Exception as e:
            logger.error(f"Error reading CSV file: {e}")
            sys.exit(1)
        
        logger.info(f"Loaded {len(records)} gas meter records from {csv_file.name}")
        if skipped_lines > 0:
            logger.warning(f"Skipped {skipped_lines} invalid lines")
        
        return records
    
    def add_gasmeter_data(self, records: List[Dict]):
        """Add gas meter data to the database."""
        # Ensure gasmeter section exists
        if 'gasmeter' not in self.database:
            self.database['gasmeter'] = {
                'records': [],
                'metadata': {
                    'description': 'Gas meter readings',
                    'unit': 'm³',
                    'last_updated': datetime.now().isoformat()
                }
            }
            logger.info("Created 'gasmeter' section in database")
        
        # Get existing records
        existing_records = self.database['gasmeter'].get('records', [])
        existing_timestamps = {r['timestamp'] for r in existing_records}
        
        # Add new records (avoid duplicates)
        new_records_count = 0
        duplicate_count = 0
        
        for record in records:
            if record['timestamp'] in existing_timestamps:
                duplicate_count += 1
                logger.debug(f"Skipping duplicate record: {record['timestamp']}")
            else:
                existing_records.append(record)
                existing_timestamps.add(record['timestamp'])
                new_records_count += 1
        
        # Sort records by timestamp
        existing_records.sort(key=lambda x: x['timestamp'])
        
        # Update database
        self.database['gasmeter']['records'] = existing_records
        self.database['gasmeter']['metadata']['last_updated'] = datetime.now().isoformat()
        self.database['gasmeter']['metadata']['total_records'] = len(existing_records)
        
        logger.info(f"Added {new_records_count} new gas meter records")
        if duplicate_count > 0:
            logger.info(f"Skipped {duplicate_count} duplicate records")
        logger.info(f"Total gas meter records in database: {len(existing_records)}")
        
        if len(existing_records) > 0:
            logger.info(f"Date range: {existing_records[0]['timestamp']} to {existing_records[-1]['timestamp']}")
    
    def process_file(self, csv_path: str):
        """Load CSV file and update database."""
        logger.info(f"Processing gas meter data from: {csv_path}")
        logger.info("=" * 60)
        
        # Load CSV data
        records = self.load_csv_file(csv_path)
        
        if not records:
            logger.error("No valid records found in CSV file")
            return
        
        # Add to database
        self.add_gasmeter_data(records)
        
        # Save database
        self._save_database()
        
        logger.info("=" * 60)
        logger.info("Gas meter data loading completed successfully!")


def print_usage():
    """Print usage information."""
    print("Gas Meter Data Loader")
    print("=" * 60)
    print("Usage: python loadGasmeterValuesIntoDatabase.py <csv_file>")
    print()
    print("CSV file format:")
    print("  Column 1: Date (MM/DD/YYYY)")
    print("  Column 2: Time (HH:MM)")
    print("  Column 3: Gas meter value (float)")
    print()
    print("Example:")
    print("  python loadGasmeterValuesIntoDatabase.py data/gasmeter_readings.csv")
    print()
    print("The script will:")
    print("  - Load gas meter readings from the CSV file")
    print("  - Add them to temperature_database.json under 'gasmeter'")
    print("  - Skip duplicate entries (same timestamp)")
    print("  - Sort records by timestamp")


def main():
    """Main function."""
    if len(sys.argv) != 2:
        print("ERROR: Missing CSV file argument\n")
        print_usage()
        sys.exit(1)
    
    csv_path = sys.argv[1]
    
    if csv_path in ['-h', '--help', 'help']:
        print_usage()
        sys.exit(0)
    
    # Process the file
    loader = GasmeterDataLoader()
    loader.process_file(csv_path)


if __name__ == "__main__":
    main()
