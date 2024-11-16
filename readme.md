## Sunshot Module

This Python script processes a dataset from an Excel file (Gyro_Sunshot.xlsx), which contains sun and gyro observations, and performs azimuth calculations based on the provided data. The output is a new Excel file (Sunshot_Output.xlsx) with calculated azimuth values and comparisons.

## Features

- **Input Data**: Reads an Excel file containing sun and gyro observations with columns for date, time, sun position (degrees, minutes, seconds), and gyro measurements.
- **Data Cleaning**: Cleans and formats the data, ensuring that all values are within valid ranges (e.g., degrees ≤ 360, minutes and seconds ≤ 60).
- **Azimuth Calculation**: Converts sun position from degrees, minutes, and seconds into decimal degrees, and calculates the azimuth using the SkyField library.
- **Output**: Outputs the processed data, including the calculated azimuths and comparisons to gyro data, in a new Excel file.

## Prerequisites

Ensure you have the following libraries installed before running the script:

- pandas
- openpyxl
- requests
- skyfield
- msvcrt (for waiting on a keypress in Windows)

You can install these libraries using pip:

``` bash
pip install pandas openpyxl requests skyfield
```

## Script Overview
- **Input File**: The script reads data from an Excel file called Gyro_Sunshot.xlsx which contains columns for the sun position (degree, minute, second) and gyro data.
- **Data Cleaning**: It checks for valid ranges for degree, minute, and second values to ensure correct data.
- **Azimuth Calculation**: It calculates the azimuth of the sun based on the timestamp of each observation using the SkyField library.
- **Output File**: The result is stored in Sunshot_Output.xlsx with the following columns:
  - S/N: Serial number
  - Datetime_UTC: Datetime in UTC
  - Obs_Sun_decimal: Sun observation in decimal degrees
  - Obs_Gyro: Gyro observation
  - Azimuth: Calculated azimuth
  - Final_Azimuth: Final azimuth after adjustment
  - C - O: Difference between Final_Azimuth and Obs_Gyro
  - Mean_C-O: Mean difference of C - O
  - Azimuth Adjustment: The final azimuth is computed using a formula, and the difference between the calculated azimuth and the observed gyro value is analyzed.

## Usage

- **Prepare the Excel File**: Ensure the Excel file Gyro_Sunshot.xlsx contains the following columns:

  - Date: The date of the observation.
  - Day, Month, Year: These will be used to create a date.
  - S/N: Serial number of each observation.
  - Time: The time of observation.
  - Obs_Sun_D, Obs_Sun_M, Obs_Sun_S: The observed sun position in degrees, minutes, and seconds.
  - Obs_Gyro: The observed gyro reading.

- **Run the Script**: Simply execute the Python script. It will:

- **Validate the input data.**
- **Calculate the azimuths.**
- **Output the results in Sunshot_Output.xlsx.**
- **Output File**: The resulting file, Sunshot_Output.xlsx, will contain the azimuth calculations, including comparisons to gyro data.

## Example

Assuming your Gyro_Sunshot.xlsx is set up as expected, running the script will output an Excel file with the calculated azimuths and comparisons.

```bash
python sunshot_azimuth_calculation.py
```

After the script runs, the output will be saved in Sunshot_Output.xlsx.

## Troubleshooting
- Ensure that the Gyro_Sunshot.xlsx file is structured correctly with the necessary columns.
- If the degree, minute, or second values are invalid (e.g., degrees > 360), the script will print an error and stop.
- The script checks for numeric values in the columns. If any column contains non-numeric data, it may throw an error.

## Key Functions
- **calc_az(tutc)**: Calculates the azimuth based on the time in UTC format (tutc). It uses the SkyField library to perform the astronomy calculations.

- **get_elevation(lat, lon)**: (Optional, commented out) This function retrieves the elevation for the given latitude and longitude using the Open-Elevation API. If enabled, it will query the elevation and use it in the azimuth calculation.


The script assumes that the data is in UTC and that the coordinates for the location are in decimal degrees (latitude and longitude).

Feel free to modify the script to fit your specific needs, especially if you're working with different data formats or want to extend the calculations!
