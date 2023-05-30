import pandas as pd
import os

# File paths
file1_path = 'PrecipData_hist.xlsx'  # Path to the first Excel file
file2_path = 'PrecipData_ssp245.xlsx'  # Path to the second Excel file
master_file_path = 'Basefile.xlsx'  # Path to the master file

# for tmax as well as tmin
file3_path = 'TMaxData_hist.xlsx'  # Path to the second Excel file
file4_path = 'TMaxData_ssp245.xlsx'  # Path to the second Excel file
file5_path = 'TMinData_hist.xlsx'  # Path to the second Excel file
file6_path = 'TMinData_ssp245.xlsx'  # Path to the second Excel file

# Read the Excel files
df1 = pd.read_excel(file1_path, header=None)
df2 = pd.read_excel(file2_path, header=None)
# Read for tmax and tmin as well
df3 = pd.read_excel(file3_path, header=None)
df4 = pd.read_excel(file4_path, header=None)
df5 = pd.read_excel(file5_path, header=None)
df6 = pd.read_excel(file6_path, header=None)


i=3
while i < 86:
    # loop through 3 to 85, all inclusive
    # Copy data from default header 'D' in each file
    data1 = df1[i].tolist()  # Column D (index 3) in file 1
    data2 = df2[i].tolist()  # Column D (index 3) in file 2
    # For tmax and tmin as well
    data3 = df3[i].tolist()
    data4 = df4[i].tolist()
    data5 = df5[i].tolist()
    data6 = df6[i].tolist()



    # Merge the data into a single column for preci
    merged_data_D = data1 + data2
    # For tmax and tmin as well
    merged_data_E = data3 + data4
    merged_data_F = data5 + data6

    # Read the master file i.e base file having year month day col till base period + ssp45
    master_df = pd.read_excel(master_file_path, header=None)

    # Add the merged data as new columns with the default headers 'D', 'E', 'F' in the master file
    master_df[3] = merged_data_D
    master_df[4] = merged_data_E
    master_df[5] = merged_data_F

    # Create a new DataFrame with the master file data and merged columns
    new_df = pd.DataFrame(master_df)

    # Set the output directory path
    output_dir = r'C:\Users\young\Desktop\NITT CE Acads\Hydrology Python Project\StationWise_hist_ssp245'
    os.makedirs(output_dir, exist_ok=True)



    # Set the new file path
    # File name has to be hist_ssp245_Station_(loop no. -2)
    filename=f"hist_ssp245_Station_{i-2}.xlsx"
    new_file_path = os.path.join(output_dir, filename)



    # Save the new DataFrame to the specified path
    new_df.to_excel(new_file_path, index=False, header=False)
    i+=1