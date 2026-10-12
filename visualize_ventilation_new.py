#!/usr/bin/env python3
"""
Simplified ventilation visualization - creates only the temperature difference heatmap.
This is the most useful visualization from STAT002 analysis.
"""

import sys
from pathlib import Path

# Add current directory to path for imports
sys.path.append(str(Path(__file__).parent))

from create_heatmap import main as create_heatmap_main

def main():
    """Run the simplified STAT002 ventilation visualization."""
    print("=" * 60)
    print("STAT002 Ventilation Analysis - Temperature Difference Heatmap")
    print("=" * 60)
    
    # Create the temperature difference heatmap (the only useful result)
    create_heatmap_main()
    
    print("\n" + "=" * 60)
    print("Visualization Complete!")
    print("=" * 60)
    print("\nOutput:")
    print("- ventilation_temperature_difference_heatmap.png")
    print("\nThis heatmap shows the temperature difference between")
    print("the room (T3_Kek) and intake (T1_BE) sensors, which provides")
    print("the most reliable indication of ventilation system behavior.")

if __name__ == "__main__":
    main()
