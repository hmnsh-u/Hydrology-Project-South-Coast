import pandas as pd
import csv
import os

def convert_xlsx_to_tab_delimited(input_file, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)
    
    # Write the DataFrame to a tab-delimited text file
    df.to_csv(output_file, sep='\t', index=False)


for i in range(1,84):
    input_file = os.path.join(r'C:\Users\young\Desktop\NITT CE Acads\Hydrology Python Project\All Stations_All Carbon Levels\StationWise_hist_ssp585', f'hist_ssp585_Station_{i}.xlsx')
    output_file = os.path.join(r'C:\Users\young\Desktop\NITT CE Acads\Hydrology Python Project\All Stations_All Carbon Levels\StationWise_hist_ssp585_txt', f'hist_ssp585_Station_{i}.txt')
    convert_xlsx_to_tab_delimited(input_file, output_file)

#to edit 1.1 to 1 in each file.

def edit_tab_delimited_value(file_path, new_value):
    # Read the tab-delimited text file
    with open(file_path, 'r', newline='') as file:
        reader = csv.reader(file, delimiter='\t')
        rows = list(reader)
    
    # Update the specific value
    rows[0][2] = new_value
    
    # Write the updated data back to the file
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        writer.writerows(rows)

# Process 10 files
for i in range(1, 84):
    file_name = f'hist_ssp585_Station_{i}.txt'
    file_path = os.path.join('C:\\Users\\young\\Desktop\\NITT CE Acads\\Hydrology Python Project\\All Stations_All Carbon Levels\\StationWise_hist_ssp585_txt', file_name)
    new_value = 1
    edit_tab_delimited_value(file_path, new_value)
