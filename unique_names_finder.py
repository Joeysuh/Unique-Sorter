import os
import re
import pandas as pd

def is_valid_folder_name(folder_name):
    
    pattern = re.compile(r'^.*_.*_longos_com_pst$') # edit this for your filter, use ".*" for a wildcard character(s)
    return bool(pattern.match(folder_name))

def get_unique_folder_names(root_directory):
    unique_names = set()

    for root, dirs, files in os.walk(root_directory):
        for folder_name in dirs:
            if is_valid_folder_name(folder_name):
                unique_names.add(folder_name)

    return list(unique_names)

def write_to_excel(unique_names, output_file):
    df = pd.DataFrame(unique_names, columns=["Folder Names"])
    df.to_excel(output_file, index=False)


root_directory = "Z:\Important Links"  # change this to your actual root directory
output_file = "Longos_Unique_Names_Test.xlsx" # change this to your desired excel workbook name

unique_folder_names = get_unique_folder_names(root_directory)
write_to_excel(unique_folder_names, output_file)

print("Sucessfully copied all unique files!")
