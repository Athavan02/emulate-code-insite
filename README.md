# **emulate_code_insite.py** - Code Description

This code performs various operations on the simulation output files from Remcom's Wireless InSite. It processes the files into excel sheets, extracts relevant data, merges multiple files, adjusts data, and creates new Excel sheets based on the processed data. The code is written in Python and utilizes the pandas and numpy libraries for data manipulation.

## Table Format of DataTX1, DataTX2, Power, and Loss Sheets

The following table provides an overview of the structure and content of the DataTX1, DataTX2, Power, and Loss sheets created by the code:

### DataTX1 Sheet

The DataTX1 sheet contains processed and adjusted data from the input files for TX1. It includes the following columns:

| Column Index | Column Name              | Description                                                                                                   |
|--------------|--------------------------|---------------------------------------------------------------------------------------------------------------|
| 1            | Path number              | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points)|
| 2            | Azimuth RX - $\phi$      | Adjusted azimuth angles (horizontal orientation) of the receivers in degrees                                  |
| 3            | Elevation RX - $\theta$  | Adjusted theta angles (vertical orientation) of receivers in degrees                                          |
| 4            | Received Power (dBm)     | Received power in dBm                                                                                         |
| 5            | Azimuth TX1 - $\phi$     | Adjusted azimuth angles (horizontal orientation) of transmitter in degrees column 				  |
| 6            | Elevation TX1 - $\theta$ | Adjusted theta angles (vertical orientation) of transmitter in degrees                                        |

### DataTX2 Sheet

The DataTX2 sheet contains processed and adjusted data from the input files for TX2. It includes the following columns:

| Column Index | Column Name              | Description                                                                                                   |
|--------------|--------------------------|---------------------------------------------------------------------------------------------------------------|
| 1            | Path number              | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points)|
| 2            | Azimuth RX - $\phi$      | Adjusted azimuth angles (horizontal orientation) of the receivers in degrees                                  |
| 3            | Elevation RX - $\theta$  | Adjusted theta angles (vertical orientation) of receivers in degrees                                          |
| 4            | Received Power (dBm)     | Received power in dBm                                                                                         |
| 5            | Azimuth TX2 - $\phi$     | Adjusted azimuth angles (horizontal orientation) of transmitter in degrees column                             |
| 6            | Elevation TX2 - $\theta$ | Adjusted theta angles (vertical orientation) of transmitter in degrees                                        |

### Power Sheet

The Power sheet contains power data extracted from the input files. It includes the following columns:

| Column Index | Column Name        | Description                                                                                                    |
|--------------|--------------------|----------------------------------------------------------------------------------------------------------------|
| 1            | Path number        | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points) |
| 2            | X (m)              | x-coordinate of receiver point                                                                                 |
| 3            | Y (m)              | y-coordinate of receiver point                                                                                 |
| 4            | Z (m)              | z-coordinate of receiver point                                                                                 |
| 5            | dist (m)           | Distance moved from first receiver point                                                                       |
| 6            | Power TX1 (dBm)    | Power data for TX1                                                                                             |
| 7            | Phase TX1 (deg)    | Phase value for TX1                                                                                            |
| 8            | Power TX2 (dBm)    | Power data for TX2                                                                                             |
| 9            | Phase TX2 (deg)    | Phase value for TX2                                                                                            |

### Loss Sheet

The Loss sheet contains loss data extracted from the input files. It includes the following columns:

| Column Index | Column Name        | Description                                                                                                    |
|--------------|--------------------|----------------------------------------------------------------------------------------------------------------|
| 1            | Path number        | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points) |
| 2            | X (m)              | x-coordinate of receiver point                                                                                 |
| 3            | Y (m)              | y-coordinate of receiver point                                                                                 |
| 4            | Z (m)              | z-coordinate of receiver point                                                                                 |
| 5            | dist (m)           | Distance moved from first receiver point                                                                       |
| 6            | FSPL TX1 (dB)      | Free space path loss for TX1                                                                                   |
| 7            | FSPL TX2 (dB)      | Free space path loss for TX2                                                                                   |
| 8            | PL TX1 (dB)        | Path loss for TX1                                                                                              |
| 9            | PL TX2 (dB)        | Path loss for TX2                                                                                              |
| 10           | XPL TX1 (dB)       | Excess path loss for TX1                                                                                       |
| 11           | XPL TX2 (dB)       | Excess path loss for TX2                                                                                       |


## How to Use

To use this code, follow these steps:

1. Ensure that Python and the necessary libraries (pandas and numpy) are installed on your system.
2. Create a new Python script and copy the code into the script.
3. Update the `folder_path` variable in the `main` function to specify the path of the folder containing the input files.
4. If needed, adjust the `substrings`, `sequence_doa_t001_02`, `sequence_dod_t001_02`, `sequence_doa_t001_01`, and `sequence_dod_t001_01` variables to match the desired file patterns.
5. If needed, adjust the `merge_output_folder`, `merge_output_file_dataTX1`, and `merge_output_file_dataTX2` variables to specify the output folder and filenames for the merged DataTX1 and DataTX2 sheets.
6. Run the script to execute the code. The processed files will be created and saved in the specified folder.

Please note that this code assumes the presence of specific file patterns and follows certain data processing logic. Make sure to review and adjust the code according to your specific requirements before running it.
