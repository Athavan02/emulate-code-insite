# **emulate_code_insite.py** - Code Description

This code performs various operations on the simulation output files from Remcom's Wireless InSite (software tool for wireless communication system analysis). It processes the files into excel sheets, extracts relevant data, merges multiple files, adjusts data, and creates new Excel sheets based on the processed data. Finally, the Excel files are converted into MATLAB files for further post-processing. The code is written in Python and utilizes the **pandas**, **scipy** and **numpy** libraries for data manipulation.

The script performs the following operations:

1. **Deleting Unwanted Files**: Removes unnecessary files created by Wireless InSite based on the provided substrings.
2. **Merging Files**: Merges specific Excel files based on the provided sequences for DOA, DOD, FSPL, PL, Power, and XPL.
3. **Adjusting Data**: Modifies specific columns in the merged data files to adjust the data according to certain offsets.
4. **Creating Loss Sheet**: Generates a Loss_ sheet by merging specific columns from the Excel files and saves it as an Excel file.
5. **Creating Power Sheet**: Generates a Power_ sheet by merging specific columns from the Excel files and saves it as an Excel file.
6. **Converting to MATLAB Files**: Converts the merged data files into MATLAB files (`.mat`) suitable for further post-processing using the iNETS ray-tracing tool.

## Prerequisites

- Python 3.x
- Required Python packages: `os`, `zipfile`, `pandas`, `numpy`, `scipy`

    ```
    pip install pandas numpy scipy
    ```

## How to Use

To use this code, follow these steps:

1. Ensure that Python and the necessary libraries (pandas, scipy and numpy) are installed on your system.
2. Create a new Python script and copy the code into the script.
3. Update the `folder_path` variable in the `main` function to specify the path of the folder containing the input files.
4. If needed, adjust the `substrings`, `sequence_doa_t001_02`, `sequence_dod_t001_02`, `sequence_doa_t001_01`, and `sequence_dod_t001_01` variables to include the desired file types to process.
5. If needed, adjust the `merge_output_folder`, `merge_output_file_dataTX1`, and `merge_output_file_dataTX2` variables to specify the output folder and filenames for the merged DataTX1 and DataTX2 sheets.
5. Specify the output folder name (`merge_output_folder`) and output file names (`merge_output_file_dataTX1`, `merge_output_file_dataTX2`, `merge_output_file_loss`, `merge_output_file_power`) according to your preference.
6. Run the script using the command `python emulate-code-insite.py`.

Please note that this code assumes the presence of specific file patterns and follows certain data processing logic. Make sure to review and adjust the code according to your specific requirements before running it.

## Table Formats of created Excel and MATLAB files (only for documentation purposes) 

The following table providess an overview of the structure and content of the created excel sheets and MATLAB files created by this python script:

### DataTX1 Sheet

The `DataTX1_.xlsx` sheet contains processed and adjusted data from the input files for TX1. It includes the following columns:

| Column Index | Column Name              | Description                                                                                                   |
|--------------|--------------------------|---------------------------------------------------------------------------------------------------------------|
| 1            | Path number              | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points)|
| 2            | Azimuth RX - $\phi$      | Adjusted azimuth angles (horizontal orientation) of the receivers in degrees [0°, 360°]                       |
| 3            | Elevation RX - $\theta$  | Adjusted theta angles (vertical orientation) of receivers in degrees [-90°, 90°]                              |
| 4            | Received Power (dBm)     | Received power in dBm                                                                                         |
| 5            | Azimuth TX1 - $\phi$     | Adjusted azimuth angles (horizontal orientation) of transmitter in degrees column [0°, 360°]	              |
| 6            | Elevation TX1 - $\theta$ | Adjusted theta angles (vertical orientation) of transmitter in degrees [-90°, 90°]                            |


### DataTX2 Sheet

The `DataTX2_.xlsx` sheet contains processed and adjusted data from the input files for TX2. It includes the following columns:

| Column Index | Column Name              | Description                                                                                                   |
|--------------|--------------------------|---------------------------------------------------------------------------------------------------------------|
| 1            | Path number              | First Walk1 (11 points), Walk2 (18 points), Walk3 (17 points) and then finally the Extra Receivers (32 points)|
| 2            | Azimuth RX - $\phi$      | Adjusted azimuth angles (horizontal orientation) of the receivers in degrees [0°, 360°]                       |
| 3            | Elevation RX - $\theta$  | Adjusted theta angles (vertical orientation) of receivers in degrees [-90°, 90°]                              |
| 4            | Received Power (dBm)     | Received power in dBm                                                                                         |
| 5            | Azimuth TX2 - $\phi$     | Adjusted azimuth angles (horizontal orientation) of transmitter in degrees column [0°, 360°]                  |
| 6            | Elevation TX2 - $\theta$ | Adjusted theta angles (vertical orientation) of transmitter in degrees [-90°, 90°]                            |


### Power Sheet

The `Power_.xlsx` sheet contains power data extracted from the input files. It includes the following columns:

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

The `Loss_.xlsx` sheet contains loss data extracted from the input files. It includes the following columns:

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


### Tx1Rx_Angles_insitefin.mat

The `Tx1Rx_Angles_insitefin.mat` file contains data related to the angles between TX1 and the receivers. The table below describes the structure of the data:

| Field Name                             | Description                                                        |
|----------------------------------------|--------------------------------------------------------------------|
| Rx_AziAngle_insite (350x350x25 double) | Contains azimuth angle (in degrees) of receiver from DataTX1_      |
| Rx_EleAngle_insite (350x350x25 double) | Contains elevation angle (in degrees) of receiver from DataTX1_    |
| Tx_AziAngle_insite (350x350x25 double) | Contains azimuth angle (in degrees) of transmitter from DataTX1_   |
| Tx_EleAngle_insite (350x350x25 double) | Contains elevation angle (in degrees) of transmitter from DataTX1_ |


### Tx2Rx_Angles_insitefin.mat

The `Tx2Rx_Angles_insitefin.mat` file contains data related to the angles between TX2 and the receivers. The table below describes the structure of the data:

| Field Name                             | Description                                                        |
|----------------------------------------|--------------------------------------------------------------------|
| Rx_AziAngle_insite (350x350x25 double) | Contains azimuth angle (in degrees) of receiver from DataTX2_      |
| Rx_EleAngle_insite (350x350x25 double) | Contains elevation angle (in degrees) of receiver from DataTX2_    |
| Tx_AziAngle_insite (350x350x25 double) | Contains azimuth angle (in degrees) of transmitter from DataTX2_   |
| Tx_EleAngle_insite (350x350x25 double) | Contains elevation angle (in degrees) of transmitter from DataTX2_ |


### SimulationRecord_insiteTX1fin.mat

The `SimulationRecord_insiteTX1fin.mat` file contains simulation record data specific to TX1. The tables below describe the structure of the two data sets in this file:

1. Receiver_Ray_insite (350x350 struct):
    
    The Receiver_Ray_insite variable is a matrix that contains information about the received rays at each receiver location. It has dimensions of 350x350, where each element of the struct corresponds to a specific receiver position and stores there its corresponding values.

| Field Name     | Description                                                                                      | 
|----------------|--------------------------------------------------------------------------------------------------|
| Power_dBm      | average power of the received rays at the specific receiver location, measured in decibels (dBm) |
| TotalPower_mW  | total power of the received rays at the specific receiver location, measured in milliwatts (mW)  |
| TotalPower_dBm | total power of the received rays at the specific receiver location, measured in decibels (dBm)   |
| Ray_count      | number of rays received at the specific receiver location                                        |
| Loss_dB        | loss of the received rays at the specific receiver location, measured in decibels (dB)           |


2. Rx_TotalPower_dBm_Matrix_insite (350x350 double):

    The Rx_TotalPower_dBm_Matrix_insite variable is a matrix that represents the received total power in decibels (dBm) at each receiver location. It has dimensions of 350x350, where each element of the matrix corresponds to a specific receiver position.


### SimulationRecord_insiteTX2fin.mat

The `SimulationRecord_insiteTX2fin.mat` file contains simulation record data specific to TX2. The tables below describe the structure of the two data sets in this file:

1. Receiver_Ray_insite (350x350 struct):
    
    The Receiver_Ray_insite variable is a matrix that contains information about the received rays at each receiver location. It has dimensions of 350x350, where each element of the struct corresponds to a specific receiver position and stores there its corresponding values.

| Field Name     | Description                                                                                      | 
|----------------|--------------------------------------------------------------------------------------------------|
| Power_dBm      | average power of the received rays at the specific receiver location, measured in decibels (dBm) |
| TotalPower_mW  | total power of the received rays at the specific receiver location, measured in milliwatts (mW)  |
| TotalPower_dBm | total power of the received rays at the specific receiver location, measured in decibels (dBm)   |
| Ray_count      | number of rays received at the specific receiver location                                        |
| Loss_dB        | loss of the received rays at the specific receiver location, measured in decibels (dB)           |


2. Rx_TotalPower_dBm_Matrix_insite (350x350 double):

    The Rx_TotalPower_dBm_Matrix_insite variable is a matrix that represents the received total power in decibels (dBm) at each receiver location. It has dimensions of 350x350, where each element of the matrix corresponds to a specific receiver position.
