import os
import zipfile
import pandas as pd
import numpy as np
from scipy.io import savemat

# This function reads a file, removes specified lines, processes the remaining lines, and saves the data to an Excel file
def process_and_save_file(file_path, lines_to_remove, output_path):
    # Read all lines from the file
    with open(file_path, 'r') as f:
        lines = f.readlines()

    # Remove the specified number of lines from the beginning (header text)
    lines = lines[lines_to_remove:]

    # Split the numbers stored as text into columns
    processed_lines = []
    for line in lines:
        columns = line.strip().split()
        # Add 'NaN' values if there are less than 2 columns
        if len(columns) < 2:
            columns = columns + ['NaN'] * (2 - len(columns))
        processed_lines.append(columns)

    # Create a DataFrame from the processed lines and save the data to an Excel file without index and header
    data = pd.DataFrame(processed_lines)
    data.to_excel(output_path, index=False, header=False)

# This function deletes unwanted files from a folder based on specified substrings
def delete_unwanted_files(folder_path, substrings):
    # Get the list of files in the specified folder
    files = os.listdir(folder_path)

    for file in files:
        file_path = os.path.join(folder_path, file)

        if '.dod.' in file or '.doa.' in file:
            lines_to_remove = 6 if '.dod.' in file or '.doa.' in file else 3
            process_and_save_file(file_path, lines_to_remove, os.path.splitext(file_path)[0] + '.xlsx')
        elif any(substring in file for substring in ['.fspl.', '.pl.', '.xpl.', '.power.']):
            process_and_save_file(file_path, 3, os.path.splitext(file_path)[0] + '.xlsx')

        # Delete the file and print the name of the deleted file
        if not any(substring in file for substring in substrings):
            os.remove(file_path)
            print(f'Deleted file: {file}')

# This function checks if a file is a valid ZIP file (otherwise opening and reading of Excel file wouldn't work)
def is_zipfile(filepath):
    try:
        with zipfile.ZipFile(filepath, 'r') as _:
            return True
    except zipfile.BadZipFile:
        return False

# This function merges multiple Excel files based on specified substrings and saves the merged data to a new Excel file
def merge_excel_files(folder_path, substrings, sequence_doa, sequence_dod, output_folder, output_file):
    # Get the list of files in the specified folder and initialize empty lists to store data frames
    files = os.listdir(folder_path)
    data_frames_doa = []
    data_frames_dod = []

    for substring in sequence_doa:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, :4]  # Keep only the first four columns
                    data_frames_doa.append(df)

    for substring in sequence_dod:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None, skiprows=1, usecols="B,C")
                    if '.dod.t001_02.' in file or '.dod.t001_01.' in file:
                        nan_row = pd.DataFrame(np.nan, index=[0], columns=range(df.shape[1]))
                        df = pd.concat([nan_row, df], ignore_index=True)
                    data_frames_dod.append(df)

    # Concatenate the data frames for each sequence and then altogether
    merged_data_doa = pd.concat(data_frames_doa, ignore_index=True)
    merged_data_dod = pd.concat(data_frames_dod, ignore_index=True)
    merged_data = pd.concat([merged_data_doa, merged_data_dod], axis=1)

    # Remove the empty column in the middle
    merged_data.dropna(axis=1, how='all', inplace=True)

    output_folder_path = os.path.join(folder_path, output_folder)
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    output_file_path = os.path.join(output_folder_path, output_file)
    merged_data.to_excel(output_file_path, index=False, header=False)

# This function adjusts the data by modifying specific columns.
def adjust_data(data):
    # Adjust the azimuth angles of the receivers in DataTX1_
    data.loc[287:311, 1] -= 31
    data.loc[417:441, 1] -= 27

    # Change the range of theta (RX and TX) from [0, 180] as given by the InSite to [-90, 90]
    data.loc[:, 2] -= 90
    data.loc[:, 5] -= 90

    # Change the range of phi (RX and TX) from [-180, 180] as given by the InSite to [0, 360]
    data.loc[data[1] <= 0, 1] = 360 + data.loc[data[1] <= 0, 1]
    data.loc[data[4] <= 0, 4] = 360 + data.loc[data[4] <= 0, 4]

    return data

# This function creates a Loss_ sheet by merging specific columns from Excel files based on provided sequences and saves it to a new Excel file
def create_loss_sheet(folder_path, substrings, sequence_fspl_t001_02, sequence_fspl_t001_01, sequence_pl_t001_02, sequence_pl_t001_01, sequence_xpl_t001_02, sequence_xpl_t001_01, output_folder, output_file):
    # Get the list of files in the specified folder and initialize empty lists to store data frames
    files = os.listdir(folder_path)
    data_frames_fspl = []
    data_frames_fspl_t001_01 = []
    data_frames_pl_t001_02 = []
    data_frames_pl_t001_01 = []
    data_frames_xpl_t001_02 = []
    data_frames_xpl_t001_01 = []

    for substring in sequence_fspl_t001_02:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, :6]  # Keep only the first six columns
                    data_frames_fspl.append(df)

    for substring in sequence_fspl_t001_01:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, 5:6]  # Keep only the 6th column
                    data_frames_fspl_t001_01.append(df)

    for substring in sequence_pl_t001_02:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, 5:6]  # Keep only the 6th column
                    data_frames_pl_t001_02.append(df)

    for substring in sequence_pl_t001_01:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, 5:6]  # Keep only the 6th column
                    data_frames_pl_t001_01.append(df)

    for substring in sequence_xpl_t001_02:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, 5:6]  # Keep only the 6th column
                    data_frames_xpl_t001_02.append(df)

    for substring in sequence_xpl_t001_01:
        for file in files:
            if substring in file:
                file_path = os.path.join(folder_path, file)
                if is_zipfile(file_path):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    df = df.iloc[:, 5:6]  # Keep only the 6th column
                    data_frames_xpl_t001_01.append(df)
    
    # Concatenate the data frames for each sequence and then altogether
    merged_data_fspl = pd.concat(data_frames_fspl, ignore_index=True)
    merged_data_fspl_t001_01 = pd.concat(data_frames_fspl_t001_01, ignore_index=True)
    merged_data_pl_t001_02 = pd.concat(data_frames_pl_t001_02, ignore_index=True)
    merged_data_pl_t001_01 = pd.concat(data_frames_pl_t001_01, ignore_index=True)
    merged_data_xpl_t001_02 = pd.concat(data_frames_xpl_t001_02, ignore_index=True)
    merged_data_xpl_t001_01 = pd.concat(data_frames_xpl_t001_01, ignore_index=True)
    merged_data = pd.concat([merged_data_fspl, merged_data_fspl_t001_01, merged_data_pl_t001_02, merged_data_pl_t001_01, merged_data_xpl_t001_02, merged_data_xpl_t001_01], axis=1)

    # Create the output folder path and save the merged data to an Excel file
    output_folder_path = os.path.join(folder_path, output_folder)
    output_file_path = os.path.join(output_folder_path, output_file)
    merged_data.to_excel(output_file_path, index=False, header=False)

# This function creates a Power_ sheet by merging specific columns from Excel files based on provided sequences and saves it to a new Excel file
def create_power_sheet(folder_path, substrings, sequence_power_t001_02, sequence_power_t001_01, output_folder, output_file):
    # Get the list of files in the specified folder and initialize an empty list to store data frames for TX1 and TX2
    files = os.listdir(folder_path)
    data_frames_power_t001_02 = []
    data_frames_power_t001_01 = []

    for sequence in [sequence_power_t001_02, sequence_power_t001_01]:
        for substring in sequence:
            for file in files:
                if substring in file:
                    file_path = os.path.join(folder_path, file)
                    if is_zipfile(file_path):
                        if sequence == sequence_power_t001_02:
                            # Read the Excel file as a data frame and append it to the list for power_t001_02
                            df = pd.read_excel(file_path, engine='openpyxl', header=None)
                            data_frames_power_t001_02.append(df)
                        elif sequence == sequence_power_t001_01:
                            # Read the Excel file as a data frame and keep only the last two columns
                            df_full = pd.read_excel(file_path, engine='openpyxl', header=None)
                            df = df_full.iloc[:, -2:]
                            data_frames_power_t001_01.append(df)

    # Concatenate the data frames 
    merged_data_power_t001_02 = pd.concat(data_frames_power_t001_02, ignore_index=True)
    merged_data_power_t001_01 = pd.concat(data_frames_power_t001_01, ignore_index=True)
    merged_data = pd.concat([merged_data_power_t001_02, merged_data_power_t001_01], axis=1)

    # Create output folder path and save merged data to an Excel file
    output_folder_path = os.path.join(folder_path, output_folder)
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    output_file_path = os.path.join(output_folder_path, output_file)
    merged_data.to_excel(output_file_path, index=False, header=False)

# This function creates finally MATLAB files (output files for further post-processing) from the Excel files in the specified folder
def ray_tracer_format(merged_data_file_path):
    # Load receiver location and the specified location of the data stored in excel sheets from 'RXMatPositions350.xlsx'
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, 'RXMatPositions350.xlsx')
    RXMatPositions350 = pd.read_excel(file_path, header=None)
    
    # ReceiverID (Name of RX), Mat_row (RX position in 350x350 matrix), Mat_col (RX position in 350x350 matrix), rows_from (begin in DataTX1_), rows_end (begin in DataTX1_), positionloss (row of Loss power)  
    RXMatPositions350.columns = ['ReceiverID', 'Mat_row', 'Mat_col', 'rows from', 'rows end', 'positionloss']
    values = RXMatPositions350.iloc[:, 1:].values

    # Load the rest of the Excel files
    file_path_dataTX1_ = os.path.join(merged_data_file_path, 'DataTX1_.xlsx')
    DataTX1_ = pd.read_excel(file_path_dataTX1_, header=None).values
    file_path_dataTX2_ = os.path.join(merged_data_file_path, 'DataTX2_.xlsx')
    DataTX2_ = pd.read_excel(file_path_dataTX2_, header=None).values
    file_path_Power_ = os.path.join(merged_data_file_path, 'Power_.xlsx')
    Power_ = pd.read_excel(file_path_Power_, header=None).values
    file_path_Loss_ = os.path.join(merged_data_file_path, 'Loss_.xlsx')
    Loss_ = pd.read_excel(file_path_Loss_, header=None).values

    # Initialize the matrices with zeros
    Tx_AziAngle_insite = np.zeros((350, 350, 25))
    Tx_EleAngle_insite = np.zeros((350, 350, 25))
    Rx_AziAngle_insite = np.zeros((350, 350, 25))
    Rx_EleAngle_insite = np.zeros((350, 350, 25))

    # Save angular values in matrices for TX1
    for ii in range(len(RXMatPositions350)):
        Rx_AziAngle_insite[values[ii,0], values[ii,1]] = DataTX1_[values[ii,2]:values[ii,3]+1,1]
        Rx_EleAngle_insite[values[ii,0], values[ii,1]] = DataTX1_[values[ii,2]:values[ii,3]+1,2]
        Tx_AziAngle_insite[values[ii,0], values[ii,1]] = DataTX1_[values[ii,2]:values[ii,3]+1,4]
        Tx_EleAngle_insite[values[ii,0], values[ii,1]] = DataTX1_[values[ii,2]:values[ii,3]+1,5] 

    # Save variables into a .mat file
    file_path_Tx1Rx_Angles = os.path.join(merged_data_file_path, 'Tx1Rx_Angles_insitefin.mat')
    savemat(file_path_Tx1Rx_Angles, {'Tx_EleAngle_insite': Tx_EleAngle_insite, 'Tx_AziAngle_insite': Tx_AziAngle_insite, 'Rx_EleAngle_insite': Rx_EleAngle_insite, 'Rx_AziAngle_insite': Rx_AziAngle_insite})

    # Save angular values in matrices for TX2
    for kk in range(len(RXMatPositions350)):
        Rx_AziAngle_insite[values[kk,0], values[kk,1]] = DataTX2_[values[kk,2]:values[kk,3]+1,1]
        Rx_EleAngle_insite[values[kk,0], values[kk,1]] = DataTX2_[values[kk,2]:values[kk,3]+1,2]
        Tx_AziAngle_insite[values[kk,0], values[kk,1]] = DataTX2_[values[kk,2]:values[kk,3]+1,4]
        Tx_EleAngle_insite[values[kk,0], values[kk,1]] = DataTX2_[values[kk,2]:values[kk,3]+1,5]

    # Save variables into a .mat file
    file_path_Tx2Rx_Angles = os.path.join(merged_data_file_path, 'Tx2Rx_Angles_insitefin.mat')
    savemat(file_path_Tx2Rx_Angles, {'Tx_EleAngle_insite': Tx_EleAngle_insite, 'Tx_AziAngle_insite': Tx_AziAngle_insite, 'Rx_EleAngle_insite': Rx_EleAngle_insite, 'Rx_AziAngle_insite': Rx_AziAngle_insite})

    # Create total power matrix
    Rx_TotalPower_dBm_Matrix_insite = np.zeros((350,350))
    for ll in range(len(RXMatPositions350)):
        Rx_TotalPower_dBm_Matrix_insite[values[ll,0]-1, values[ll,1]-1] = Power_[values[ll,4]-1,6]

    # Initialize power matrix with zeros    
    Receiver_Ray_insite = np.zeros((350,350), dtype={'names':('Power_dBm','TotalPower_mW','TotalPower_dBm','Ray_count','Loss_dB'),'formats':('f8','f8','f8','i4','f8')})
                                   
    # Save power values for TX1
    for pp in range(len(RXMatPositions350)):
        Receiver_Ray_insite[values[pp,0]-1, values[pp,1]-1]['Power_dBm'] = np.mean(DataTX1_[values[pp,2]:values[pp,3]+1,3])
        Receiver_Ray_insite[values[pp,0]-1, values[pp,1]-1]['TotalPower_mW'] = 10**((abs(Power_[values[pp,4]-1,7])-30)/10)
        Receiver_Ray_insite[values[pp,0]-1, values[pp,1]-1]['TotalPower_dBm'] = Power_[values[pp,4]-1,7]
        Receiver_Ray_insite[values[pp,0]-1, values[pp,1]-1]['Ray_count'] = values[pp,3]-values[pp,2]+1
        Receiver_Ray_insite[values[pp,0]-1, values[pp,1]-1]['Loss_dB'] = Loss_[values[pp,4]-1,7]

    # Save variables into a .mat file
    file_path_SimRecTX1 = os.path.join(merged_data_file_path, 'SimulationRecord_insiteTX1fin.mat')
    savemat(file_path_SimRecTX1, {'Receiver_Ray_insite': Receiver_Ray_insite, 'Rx_TotalPower_dBm_Matrix_insite': Rx_TotalPower_dBm_Matrix_insite})

    # Save power values for TX2
    for hh in range(len(RXMatPositions350)):
        Receiver_Ray_insite[values[hh,0]-1, values[hh,1]-1]['Power_dBm'] = np.mean(DataTX2_[values[hh,2]:values[hh,3]+1,3])
        Receiver_Ray_insite[values[hh,0]-1, values[hh,1]-1]['TotalPower_mW'] = 10**((abs(Power_[values[hh,4]-1,7])-30)/10)
        Receiver_Ray_insite[values[hh,0]-1, values[hh,1]-1]['TotalPower_dBm'] = Power_[values[hh,4]-1,7]
        Receiver_Ray_insite[values[hh,0]-1, values[hh,1]-1]['Ray_count'] = values[hh,3]-values[hh,2]+1
        Receiver_Ray_insite[values[hh,0]-1, values[hh,1]-1]['Loss_dB'] = Loss_[values[hh,4]-1,8]

    # Save variables into a .mat file
    file_path_SimRecTX2 = os.path.join(merged_data_file_path, 'SimulationRecord_insiteTX2fin.mat')
    savemat(file_path_SimRecTX2, {'Receiver_Ray_insite': Receiver_Ray_insite, 'Rx_TotalPower_dBm_Matrix_insite': Rx_TotalPower_dBm_Matrix_insite})

def main():
    # Specify location for saving and what kind of data you want to store (DOA, DOD, FSPL, PL, Power, XPL or other types from Wireless InSite)
    folder_path = r'C:\Users\Athavan\Desktop\Code\emulate-code-insite\InSiteOutput'
    substrings = ['.dod.', '.doa.', '.fspl.', '.pl.', '.power.', '.xpl.']

    sequence_doa_t001_02 = ['.doa.t001_02.r010', '.doa.t001_02.r009', '.doa.t001_02.r007', '.doa.t001_02.r011']
    sequence_dod_t001_02 = ['.dod.t001_02.r010', '.dod.t001_02.r009', '.dod.t001_02.r007', '.dod.t001_02.r011']
    sequence_doa_t001_01 = ['.doa.t001_01.r010', '.doa.t001_01.r009', '.doa.t001_01.r007', '.doa.t001_01.r011']
    sequence_dod_t001_01 = ['.dod.t001_01.r010', '.dod.t001_01.r009', '.dod.t001_01.r007', '.dod.t001_01.r011']
    merge_output_folder = 'merged_data'
    merge_output_file_dataTX1 = 'DataTX1_.xlsx'
    merge_output_file_dataTX2 = 'DataTX2_.xlsx'

    # Delete unnecessary files created by Wireless InSite from folder
    delete_unwanted_files(folder_path, substrings)
    merge_excel_files(folder_path, substrings, sequence_doa_t001_02, sequence_dod_t001_02, merge_output_folder, merge_output_file_dataTX1)
    merge_excel_files(folder_path, substrings, sequence_doa_t001_01, sequence_dod_t001_01, merge_output_folder, merge_output_file_dataTX2)

    # Read the merged files, adjust the data by correcting some offsets, and save the adjusted data
    output_file_path1 = os.path.join(folder_path, merge_output_folder, merge_output_file_dataTX1)
    output_file_path2 = os.path.join(folder_path, merge_output_folder, merge_output_file_dataTX2)
    dataTX1_fin = pd.read_excel(output_file_path1, header=None, engine='openpyxl')
    dataTX2_fin = pd.read_excel(output_file_path2, header=None, engine='openpyxl')
    adjusted_dataTX1_fin = adjust_data(dataTX1_fin)
    adjusted_dataTX2_fin = adjust_data(dataTX2_fin)
    adjusted_dataTX1_fin.to_excel(output_file_path1, index=False, header=False)
    adjusted_dataTX2_fin.to_excel(output_file_path2, index=False, header=False)

    # Create the Loss_ sheet
    sequence_fspl_t001_02 = ['.fspl.t001_02.r010', '.fspl.t001_02.r009', '.fspl.t001_02.r007', '.fspl.t001_02.r011']
    sequence_fspl_t001_01 = ['.fspl.t001_01.r010', '.fspl.t001_01.r009', '.fspl.t001_01.r007', '.fspl.t001_01.r011']
    sequence_pl_t001_02 = ['.pl.t001_02.r010', '.pl.t001_02.r009', '.pl.t001_02.r007', '.pl.t001_02.r011']
    sequence_pl_t001_01 = ['.pl.t001_01.r010', '.pl.t001_01.r009', '.pl.t001_01.r007', '.pl.t001_01.r011']
    sequence_xpl_t001_02 = ['.xpl.t001_02.r010', '.xpl.t001_02.r009', '.xpl.t001_02.r007', '.xpl.t001_02.r011']
    sequence_xpl_t001_01 = ['.xpl.t001_01.r010', '.xpl.t001_01.r009', '.xpl.t001_01.r007', '.xpl.t001_01.r011']
    merge_output_file_loss = 'Loss_.xlsx'
    create_loss_sheet(folder_path, substrings, sequence_fspl_t001_02, sequence_fspl_t001_01, sequence_pl_t001_02, sequence_pl_t001_01, sequence_xpl_t001_02, sequence_xpl_t001_01, merge_output_folder, merge_output_file_loss)

    # Create the Power_ sheet
    sequence_power_t001_02 = ['.power.t001_02.r010', '.power.t001_02.r009', '.power.t001_02.r007', '.power.t001_02.r011']
    sequence_power_t001_01 = ['.power.t001_01.r010', '.power.t001_01.r009', '.power.t001_01.r007', '.power.t001_01.r011']
    merge_output_file_power = 'Power_.xlsx'
    create_power_sheet(folder_path, substrings, sequence_power_t001_02, sequence_power_t001_01, merge_output_folder, merge_output_file_power)

    # Change the Excel files into iNETS ray-tracing compatible format (MATLAB files)
    merged_data_file_path = os.path.join(folder_path, merge_output_folder)
    ray_tracer_format(merged_data_file_path)

if __name__ == '__main__':
    main()