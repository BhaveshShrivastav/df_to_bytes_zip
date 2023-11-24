import os
import pandas as pd
import numpy as np
import shutil
import os
from io import BytesIO
import zipfile

def create_directory(directory_name, location) :
    #Get the current working directory
    cwd = os.getcwd()
    print("directory:", cwd)

    # Create the full path for the new directory
    full_path = os.path.join(cwd, location, directory_name)

    # Check if the directory already exists
    if not os.path.exists(full_path):
        # Create the directory
        os.makedirs(full_path)
        print(f"Directory '{directory_name}' created at location: {full_path}")
        return full_path
    else:
        print(f"Directory '{directory_name}' already exists at location: {full_path}")
        return full_path


def create_df(number: int) -> list():
    # Number of DataFrames to create
    # if not number > 0 :
    #     num_dataframes = 5
    # else:
    #     num_dataframes = number
    num_dataframes = number
    # List to store the DataFrames
    dataframes_list = []

    # Loop to create DataFrames
    for i in range(num_dataframes):
        # Create a DataFrame with dummy data
        df = pd.DataFrame({
            'Column1': np.random.rand(10),
            'Column2': np.random.randint(0, 100, 10),
            'Column3': np.random.choice(['A', 'B', 'C'], 10)
        })

        # Append the DataFrame to the list
        dataframes_list.append(df)

    # Accessing individual DataFrames in the list
    for idx, df in enumerate(dataframes_list):
        print(f"DataFrame {idx + 1}:")
        print(df)
        print("\n")
    return dataframes_list


def write_excel_from_dataframes(file_path, dataframes_list, sheet_names=None):
    """
    Write a list of DataFrames to an Excel file with each DataFrame in a different sheet.

    Parameters:
        - file_path (str): The file path for the Excel file.
        - dataframes_list (list): A list of Pandas DataFrames to be written to the Excel file.
        - sheet_names (list or None): A list of sheet names corresponding to each DataFrame.
          If None, default sheet names (Sheet1, Sheet2, ...) will be used.

    Returns:
        None
    """
    if not sheet_names:
        sheet_names = [f"Sheet{i+1}" for i in range(len(dataframes_list))]

    # Create ExcelWriter object
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        # Write each DataFrame to a different sheet
        for df, sheet_name in zip(dataframes_list, sheet_names):
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Excel file '{file_path}' created successfully.")



# Function to create a zip archive of a directory
def create_zip_bytes(directory_path):
    # Create a BytesIO object to store the zip file in memory
    zip_buffer = BytesIO()

    # Create a ZipFile object with write mode
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        # Iterate over all the files in the directory and add them to the zip file
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, directory_path)
                zip_file.write(file_path, arcname=arcname)

    # Seek to the beginning of the BytesIO object
    zip_buffer.seek(0)

    return zip_buffer



# process
def process():
    # Example usage
    directory_name = "response"
    location = "dummy"
    for i in range(1,10):
        path = create_directory(directory_name, location)
        dataframes_list=create_df(5)
        # print(dataframes_list)
        excel_file_path = os.path.join(path, 'File_'+str(i)+'_output_file_multiple.xlsx')
        # Call the function to write DataFrames to Excel file
        write_excel_from_dataframes(excel_file_path, dataframes_list)
        print("here is the directory:",path)
        zip_bytes_data = create_zip_bytes(path)
    return zip_bytes_data

# Function to save zip bytes to a file
def save_zip_bytes(zip_bytes, output_directory, zip_filename):
    zip_file_path = os.path.join(output_directory, zip_filename)

    with open(zip_file_path, 'wb') as zip_file:
        zip_file.write(zip_bytes.read())

    print(f"Zip file saved to: {zip_file_path}")

def delete_directory(directory_path):
    try:
        # Use shutil.rmtree to delete the directory and its contents
        shutil.rmtree(directory_path)
        print(f"Directory '{directory_path}' successfully deleted.")
    except OSError as e:
        print(f"Error: {e}")

def main():
    zip_bytes=process()
    save_zip_bytes(zip_bytes=zip_bytes, output_directory='./results/', zip_filename='outputfile.zip')
    delete_directory('./dummy/response')
    

main()
