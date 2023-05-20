import os
import zipfile
import pandas as pd
import numpy as np

def process_and_save_file(file_path, lines_to_remove, output_path):
    with open(file_path, 'r') as f:
        lines = f.readlines()

    lines = lines[lines_to_remove:]

    processed_lines = []
    for line in lines:
        columns = line.strip().split()
        if len(columns) < 2:
            columns = columns + ['NaN'] * (2 - len(columns))
        processed_lines.append(columns)

    data = pd.DataFrame(processed_lines)
    data.to_excel(output_path, index=False, header=False)

def delete_unwanted_files(folder_path, substrings):
    files = os.listdir(folder_path)

    for file in files:
        file_path = os.path.join(folder_path, file)

        if '.dod.' in file or '.doa.' in file:
            lines_to_remove = 6 if '.dod.' in file or '.doa.' in file else 3
            process_and_save_file(file_path, lines_to_remove, os.path.splitext(file_path)[0] + '.xlsx')
        elif any(substring in file for substring in ['.fspl.', '.pl.', '.xpl.', '.power.']):
            process_and_save_file(file_path, 3, os.path.splitext(file_path)[0] + '.xlsx')

        if not any(substring in file for substring in substrings):
            os.remove(file_path)
            print(f'Deleted file: {file}')

def is_zipfile(filepath):
    try:
        with zipfile.ZipFile(filepath, 'r') as _:
            return True
    except zipfile.BadZipFile:
        return False

def merge_excel_files(folder_path, substrings, sequence_doa, sequence_dod, output_folder, output_file):
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

def create_loss_sheet(folder_path, substrings, sequence_fspl_t001_02, sequence_fspl_t001_01, sequence_pl_t001_02, sequence_pl_t001_01, sequence_xpl_t001_02, sequence_xpl_t001_01, output_folder, output_file):
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

    # Add the following three for loops to collect the specified data
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

    merged_data_fspl = pd.concat(data_frames_fspl, ignore_index=True)
    merged_data_fspl_t001_01 = pd.concat(data_frames_fspl_t001_01, ignore_index=True)
    merged_data_pl_t001_02 = pd.concat(data_frames_pl_t001_02, ignore_index=True)
    merged_data_pl_t001_01 = pd.concat(data_frames_pl_t001_01, ignore_index=True)
    merged_data_xpl_t001_02 = pd.concat(data_frames_xpl_t001_02, ignore_index=True)
    merged_data_xpl_t001_01 = pd.concat(data_frames_xpl_t001_01, ignore_index=True)

    merged_data = pd.concat([merged_data_fspl, merged_data_fspl_t001_01, merged_data_pl_t001_02, merged_data_pl_t001_01, merged_data_xpl_t001_02, merged_data_xpl_t001_01], axis=1)

    output_folder_path = os.path.join(folder_path, output_folder)
    output_file_path = os.path.join(output_folder_path, output_file)
    merged_data.to_excel(output_file_path, index=False, header=False)

def create_power_sheet(folder_path, substrings, sequence_power_t001_02, sequence_power_t001_01, output_folder, output_file):
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
                            df = pd.read_excel(file_path, engine='openpyxl', header=None)
                            data_frames_power_t001_02.append(df)
                        elif sequence == sequence_power_t001_01:
                            df_full = pd.read_excel(file_path, engine='openpyxl', header=None)
                            df = df_full.iloc[:, -2:]
                            data_frames_power_t001_01.append(df)

    merged_data_power_t001_02 = pd.concat(data_frames_power_t001_02, ignore_index=True)
    merged_data_power_t001_01 = pd.concat(data_frames_power_t001_01, ignore_index=True)
    merged_data = pd.concat([merged_data_power_t001_02, merged_data_power_t001_01], axis=1)

    output_folder_path = os.path.join(folder_path, output_folder)
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    output_file_path = os.path.join(output_folder_path, output_file)
    merged_data.to_excel(output_file_path, index=False, header=False)

def main():
    folder_path = r'C:\Users\Athavan\Desktop\InSiteOutput'
    substrings = ['.dod.', '.doa.', '.fspl.', '.pl.', '.power.', '.xpl.']
    sequence_doa_t001_02 = ['.doa.t001_02.r010', '.doa.t001_02.r009', '.doa.t001_02.r007', '.doa.t001_02.r011']
    sequence_dod_t001_02 = ['.dod.t001_02.r010', '.dod.t001_02.r009', '.dod.t001_02.r007', '.dod.t001_02.r011']
    sequence_doa_t001_01 = ['.doa.t001_01.r010', '.doa.t001_01.r009', '.doa.t001_01.r007', '.doa.t001_01.r011']
    sequence_dod_t001_01 = ['.dod.t001_01.r010', '.dod.t001_01.r009', '.dod.t001_01.r007', '.dod.t001_01.r011']
    merge_output_folder = 'merged_data'
    merge_output_file_dataTX1 = 'DataTX1_.xlsx'
    merge_output_file_dataTX2 = 'DataTX2_.xlsx'

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

if __name__ == '__main__':
    main()