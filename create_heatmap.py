#!/usr/bin/env python3
"""
Create temperature difference heatmap visualization for ventilation analysis.
Shows the temperature difference between T3_Kek (room) and T1_BE (intake).
"""

import json
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

def load_ventilation_data():
    """Load STAT002 results from JSON file."""
    results_path = Path("output/STAT002_ventilation_analysis.json")
    
    if not results_path.exists():
        print(f"Error: {results_path} not found. Please run temperature_statistics.py first.")
        return None
    
    with open(results_path, 'r') as f:
        data = json.load(f)
    
    return data

def create_temperature_difference_heatmap(data, output_path="output/ventilation_temperature_difference_heatmap.png"):
    """
    Create a heatmap showing temperature difference between T3_Kek (room) and T1_BE (intake).
    
    Args:
        data: STAT002 analysis results
        output_path: Path to save the temperature difference heatmap image
    """
    results = data['detailed_results']
    
    # Convert to DataFrame
    df = pd.DataFrame(results)
    df['timestamp'] = pd.to_datetime(df['timestamp'], format='ISO8601')
    
    # Calculate temperature difference (Room - Intake)
    df['temp_difference'] = df['T3_Kek'] - df['T1_BE']
    
    # Extract date and hour
    df['date'] = df['timestamp'].dt.date
    df['hour'] = df['timestamp'].dt.hour
    
    # Get date range
    start_date = df['date'].min()
    end_date = df['date'].max()
    
    # Create a complete date range
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # Initialize the temperature difference matrix (days x 24 hours)
    temp_diff_matrix = np.full((len(date_range), 24), np.nan)
    
    # Fill the matrix with temperature difference values
    for _, row in df.iterrows():
        date_idx = (row['date'] - start_date).days
        hour_idx = row['hour']
        
        if 0 <= date_idx < len(date_range) and 0 <= hour_idx < 24:
            temp_diff_matrix[date_idx, hour_idx] = row['temp_difference']
    
    # Create the visualization
    fig, ax = plt.subplots(figsize=(20, 12))
    
    # Calculate color scale limits for better visualization
    # Use percentiles to avoid extreme outliers affecting the scale
    valid_data = temp_diff_matrix[~np.isnan(temp_diff_matrix)]
    if len(valid_data) > 0:
        vmin = np.percentile(valid_data, 1)  # 1st percentile
        vmax = np.percentile(valid_data, 99)  # 99th percentile
        # Make scale symmetric around zero for better interpretation
        abs_max = max(abs(vmin), abs(vmax))
        vmin, vmax = -abs_max, abs_max
    else:
        vmin, vmax = -5, 5  # Default range
    
    # Create the heatmap with thermal colormap
    im = ax.imshow(temp_diff_matrix, cmap='RdBu_r', vmin=vmin, vmax=vmax,
                   aspect='auto', origin='upper')
    
    # Configure x-axis (hours)
    ax.set_xlim(-0.5, 23.5)
    ax.set_xticks(range(0, 24, 2))
    ax.set_xticklabels([f"{h:02d}:00" for h in range(0, 24, 2)])
    ax.set_xlabel('Hour of Day')
    
    # Configure y-axis (dates)
    date_ticks = range(0, len(date_range), 7)
    ax.set_yticks(date_ticks)
    ax.set_yticklabels([date_range[i].strftime('%Y-%m-%d') for i in date_ticks])
    ax.set_ylabel('Date')
    
    # Add grid
    ax.set_xticks(np.arange(-0.5, 24, 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(date_range), 1), minor=True)
    ax.grid(which='minor', color='white', linestyle='-', linewidth=0.5, alpha=0.3)
    
    # Add colorbar
    cbar = plt.colorbar(im, ax=ax)
    cbar.set_label('Temperature Difference (T3_Kek - T1_BE) [°C]', rotation=270, labelpad=20)
    
    # Add title with interpretation guide
    ax.set_title('Temperature Difference: Room (T3_Kek) - Intake (T1_BE)\n'
                 f'From {start_date} to {end_date}\n'
                 'Red = Room warmer than intake | Blue = Intake warmer than room',
                 fontsize=14, pad=20)
    
    # Add interpretation text
    interpretation_text = (
        "Interpretation:\n"
        "• Red areas: Room temperature > Intake temperature (likely ventilation OFF)\n"
        "• Blue areas: Intake temperature > Room temperature (possible ventilation ON or external influence)\n"
        "• White areas: Similar temperatures between room and intake"
    )
    
    ax.text(0.02, 0.98, interpretation_text, transform=ax.transAxes,
            ha='left', va='top', fontsize=10,
            bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', alpha=0.8))
    
    # Add statistics
    if len(valid_data) > 0:
        stats_text = (f"Statistics: Mean: {np.mean(valid_data):.1f}°C | "
                     f"Std: {np.std(valid_data):.1f}°C | "
                     f"Range: {np.min(valid_data):.1f}°C to {np.max(valid_data):.1f}°C")
        
        ax.text(0.5, -0.08, stats_text, transform=ax.transAxes,
                ha='center', va='top', fontsize=10,
                bbox=dict(boxstyle='round,pad=0.5', facecolor='lightblue', alpha=0.7))
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    print(f"Temperature difference heatmap saved to: {output_path}")
    
    return fig

def main():
    """Main function to create temperature difference heatmap."""
    print("Creating Temperature Difference Heatmap...")
    
    # Load data
    data = load_ventilation_data()
    if data is None:
        return
    
    print(f"\nLoaded {data['total_data_points']:,} data points")
    print(f"Time range: {data['time_range']['start']} to {data['time_range']['end']}")
    
    # Create temperature difference heatmap
    print("\nCreating temperature difference heatmap...")
    create_temperature_difference_heatmap(data)
    
    print("\nVisualization completed!")
    print("Generated file:")
    print("- ventilation_temperature_difference_heatmap.png: Room-Intake temperature difference")

if __name__ == "__main__":
    main()
