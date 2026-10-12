# Overview

The generate_celendars.py script creates "calendar images" for given time-value series. The time series contains timestamp-value pairs and the calendar image is a point matrix where rows represent the days (vertical axis is time in days), the columns represent samples inside the day (usually 5 minute is the sampling time), so horizontal axis is the time inside the day, and the pixel values represent the value of the time series using a heatmap color scheme with normalized values so that the whole dataset uses the whole range of the heatmap. The file format is PNG and it should also contain axis labels and values to interpret the heatmap (figure legend).

# Calendar images

The following calendar images are created (with code names in brackets which can be used as a filename.) The images should cover the whole time range of the data. Assume that all data have a sample interval of exactly 5 minutes. Display actual time labels on the horizontal axis like "08:00".

Error handling:

## [TempCal_Z1] Temperature measured by the device T8_Z1

Temperature values measured by the device T8_Z1, taken from temperature_database.json . Missing values should be indicated by white pixels. Use the "viridis" colormap.

## [Heating_Z1] Temperature measured by the device T8_Z1

Acitivity of the heating in zone 1, taken from `heating_cycles.json` for device T8_Z1. The values in the image should be binary.Use red for active heating and white for passive. Missing data should be gray.

## [TempCal_Z2] and [Heating_Z2]

Same as the ones for Z1 using device T6_Z2.

## [TempDiff] Temperature difference between devices T3_Kek and T2_Terasz

Temperature values for the devices T3_Kek and T2_Terasz should be taken from temperature_database.json . Missing values should be indicated by white pixels. Use the "viridis" colormap.

## [Temp_Outside] Temperature of T2_Terasz

