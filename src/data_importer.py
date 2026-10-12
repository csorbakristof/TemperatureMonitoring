"""
Data Import Module (IN001) - Importing new data

This module handles importing temperature data from multiple ZIP files into a central JSON database.
It uses the TemperatureDataProcessor for CSV processing and focuses on database operations.
"""

import json
import os
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Set, Tuple, Any

# Add src directory to path for imports
sys.path.append(str(Path(__file__).parent))

from temperature_processor import TemperatureDataProcessor

logger = logging.getLogger(__name__)


class TemperatureDataImporter:
    """Handles importing and merging temperature data from multiple ZIP files."""
    
    def __init__(self, json_db_path: str = "data/temperature_database.json"):
        self.json_db_path = Path(json_db_path)
        self.processor = TemperatureDataProcessor()
        self.database: Dict = self._load_database()
        
    def _load_database(self) -> Dict:
        """Load existing JSON database or create new empty one."""
        if self.json_db_path.exists():
            try:
                with open(self.json_db_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    logger.info(f"Loaded existing database with {len(data.get('devices', {}))} devices")
                    return data
            except (json.JSONDecodeError, FileNotFoundError) as e:
                logger.warning(f"Could not load existing database: {e}. Creating new one.")
        
        # Create new database structure
        return {
            "metadata": {
                "created": datetime.now().isoformat(),
                "last_updated": datetime.now().isoformat(),
                "version": "1.0.0",
                "total_records": 0
            },
            "devices": {},
            "import_history": []
        }
    
    def _save_database(self) -> None:
        """Save database to JSON file."""
        self.database["metadata"]["last_updated"] = datetime.now().isoformat()
        
        # Ensure directory exists
        self.json_db_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Save with pretty formatting
        with open(self.json_db_path, 'w', encoding='utf-8') as f:
            json.dump(self.database, f, indent=2, default=str, ensure_ascii=False)
        
        logger.info(f"Database saved to {self.json_db_path}")
    
    def _process_device_data(self, device_data: Dict) -> Tuple[int, int]:
        """
        Process a single device's data and merge into database.
        
        Returns:
            Tuple of (new_records_added, duplicate_records_skipped)
        """
        device_name = device_data['device_name']
        records = device_data['data']
        
        if not records:
            logger.warning(f"No data records found for device {device_name}")
            return 0, 0
        
        # Initialize device in database if not exists
        if device_name not in self.database['devices']:
            self.database['devices'][device_name] = {
                "device_name": device_name,
                "first_seen": datetime.now().isoformat(),
                "records": [],
                "existing_records": set()
            }
        
        device_db = self.database['devices'][device_name]
        existing_records = set()
        
        # Build set of existing records for duplicate detection
        for existing_record in device_db['records']:
            record_key = (
                existing_record['timestamp'],
                existing_record['temperature'],
                existing_record['humidity'],
                existing_record['battery_mv']
            )
            existing_records.add(record_key)
        
        new_records = 0
        duplicates = 0
        
        for record in records:
            # Convert datetime to ISO string for JSON storage
            timestamp_str = record['timestamp'].isoformat()
            
            # Create record key for duplicate detection (in-memory comparison)
            record_key = (
                timestamp_str,
                record['temperature'],
                record['humidity'],
                record['battery_mv']
            )
            
            if record_key not in existing_records:
                # New record - add to database
                json_record = {
                    "timestamp": timestamp_str,
                    "temperature": record['temperature'],
                    "humidity": record['humidity'],
                    "battery_mv": record['battery_mv']
                }
                
                device_db['records'].append(json_record)
                existing_records.add(record_key)
                new_records += 1
            else:
                duplicates += 1
        
        # Update device metadata
        device_db['total_records'] = len(device_db['records'])
        device_db['last_updated'] = datetime.now().isoformat()
        
        # Sort records by timestamp for better organization
        device_db['records'].sort(key=lambda x: x['timestamp'])
        
        logger.info(f"Device {device_name}: {new_records} new records, {duplicates} duplicates skipped")
        return new_records, duplicates
    
    def import_zip_files(self, data_folder: str = "data") -> Dict[str, Any]:
        """
        Import all TempLogs*.zip files from the specified folder.
        
        Args:
            data_folder: Path to folder containing ZIP files
            
        Returns:
            Dictionary with import statistics
        """
        data_path = Path(data_folder)
        
        # Find all TempLogs*.zip files
        zip_files = list(data_path.glob("TempLogs*.zip"))
        
        if not zip_files:
            logger.warning(f"No TempLogs*.zip files found in {data_folder}")
            return {"error": "No ZIP files found"}
        
        logger.info(f"Found {len(zip_files)} ZIP files to process")
        
        import_stats = {
            "start_time": datetime.now().isoformat(),
            "zip_files_processed": 0,
            "devices_found": 0,
            "total_new_records": 0,
            "total_duplicates": 0,
            "files_processed": [],
            "errors": []
        }
        
        for zip_file in sorted(zip_files):
            try:
                logger.info(f"Processing {zip_file.name}...")
                
                # Process ZIP file
                all_device_data = self.processor.process_zip_file(str(zip_file))
                
                if not all_device_data:
                    logger.warning(f"No device data found in {zip_file.name}")
                    continue
                
                file_stats = {
                    "filename": zip_file.name,
                    "devices": [],
                    "new_records": 0,
                    "duplicates": 0
                }
                
                # Process each device's data
                for device_data in all_device_data:
                    new_records, duplicates = self._process_device_data(device_data)
                    
                    file_stats["devices"].append({
                        "name": device_data['device_name'],
                        "new_records": new_records,
                        "duplicates": duplicates
                    })
                    
                    file_stats["new_records"] += new_records
                    file_stats["duplicates"] += duplicates
                
                import_stats["files_processed"].append(file_stats)
                import_stats["zip_files_processed"] += 1
                import_stats["total_new_records"] += file_stats["new_records"]
                import_stats["total_duplicates"] += file_stats["duplicates"]
                
            except Exception as e:
                error_msg = f"Error processing {zip_file.name}: {str(e)}"
                logger.error(error_msg)
                import_stats["errors"].append(error_msg)
        
        # Update database metadata
        import_stats["devices_found"] = len(self.database['devices'])
        import_stats["end_time"] = datetime.now().isoformat()
        
        # Update total records count
        total_records = sum(len(device['records']) for device in self.database['devices'].values())
        self.database["metadata"]["total_records"] = total_records
        
        # Add import to history
        self.database["import_history"].append(import_stats)
        
        # Save updated database
        self._save_database()
        
        logger.info(f"Import completed: {import_stats['total_new_records']} new records, "
                   f"{import_stats['total_duplicates']} duplicates from {import_stats['zip_files_processed']} files")
        
        return import_stats
    
    def get_database_summary(self) -> Dict[str, Any]:
        """Get summary information about the current database."""
        if not self.database['devices']:
            return {"message": "Database is empty"}
        
        summary = {
            "total_devices": len(self.database['devices']),
            "total_records": self.database["metadata"]["total_records"],
            "last_updated": self.database["metadata"]["last_updated"],
            "devices": []
        }
        
        for device_name, device_data in self.database['devices'].items():
            device_summary = {
                "name": device_name,
                "total_records": len(device_data['records']),
                "first_record": None,
                "last_record": None
            }
            
            if device_data['records']:
                device_summary["first_record"] = device_data['records'][0]['timestamp']
                device_summary["last_record"] = device_data['records'][-1]['timestamp']
            
            summary["devices"].append(device_summary)
        
        return summary


def main():
    """Main function for data import."""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    importer = TemperatureDataImporter()
    
    print("Temperature Data Importer (IN001)")
    print("=" * 50)
    
    # Import all ZIP files
    import_stats = importer.import_zip_files()
    
    # Print results
    if "error" in import_stats:
        print(f"Error: {import_stats['error']}")
        return
    
    print(f"\nImport Results:")
    print(f"ZIP files processed: {import_stats['zip_files_processed']}")
    print(f"Devices found: {import_stats['devices_found']}")
    print(f"New records added: {import_stats['total_new_records']}")
    print(f"Duplicate records skipped: {import_stats['total_duplicates']}")
    
    if import_stats['errors']:
        print(f"\nErrors encountered:")
        for error in import_stats['errors']:
            print(f"  - {error}")
    
    # Show database summary
    print(f"\nDatabase Summary:")
    summary = importer.get_database_summary()
    print(f"Total devices: {summary['total_devices']}")
    print(f"Total records: {summary['total_records']}")
    
    for device in summary['devices']:
        print(f"  {device['name']}: {device['total_records']} records "
              f"({device['first_record']} to {device['last_record']})")


if __name__ == "__main__":
    main()
