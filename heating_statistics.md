# Overview

The heating_statistics.py script is analysing the heating patterns and their relationships to outside temperature and gas consumption.

Only plot for time intervals where all necessary data are available. For heating cycle data, use `heating_cycles.json` or print an error message if not available and ask the user to run `detect_heating.py`.

Output filenames should match the code name of the feature in brackets in the headings. For the diagrams, use png format. Also export the data into CSV format with the same name but csv extension. Write everything in the `output` directory.

Also export a summary statistic into `heating_statistics.txt`.

## Gas meter input data

Gas meter values are available in the temperature_database.json under `gasmeter`. If it is not available, ask the user to run `loadGasmeterValuesIntoDatabase.py`.

# Outputs

## [MeanDiff_And_HeatingCycleCount_Plot] Mean difference of internal and external temperature, and the number of daily heat cycles

This is a diagram with time (days) on the horizontal axis and two datasets as vertical axis.
Left vertical axis (blue color): daily mean of the difference between external and internal temperature. Temperature values come from temperature_database.json, device T3_Kek is internal temperature, T2_Terasz is external temperature. Difference should be shown using blue dots without connecting lines (in matplotlib, "."). Use the formula `mean(T3_Kek - T2_Terasz)`.
Right vertical axis (red color): daily number of heating cycles for zone 2 (device T6_Z2).

## [MeanDiff_And_HeatingCycleCount_XY] Similar to MeanDiff_And_HeatingCycleCount_Plot, but with XY Plot

This is a diagram has the same data as MeanDiff_And_HeatingCycleCount_Plot, but the daily mean temperature difference is on the horizontal axis and the number of daily heat cycles is on the vertical axis, so this is an X-Y plot. Dots should not be connected.

## [GasVsCycleCount] Gas consumption vs heating cycle count

This is an X-Y scatter plot, one dot for every gas meter measurement.
X-axis: gasmeter value change since last measurement (difference)
Y-axis: the total number of heating cycles since the last gasmeter measurement
