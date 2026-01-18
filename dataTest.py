import os
import pandas as pd


user_profile = os.environ["USERPROFILE"]
data_file_folder = os.path.join(user_profile, "OneDrive", "Documents", "projects", "python zero to hero", "data"
                                )
print(repr(data_file_folder))
print(os.path.isdir(data_file_folder))

print(f"checking:  {data_file_folder}")

# creating a list to hold the dataframe
all_dataframes = []

# fault tolerance
# check if folder exist, if so, then forloop if file find it is xlsx files, concate
if os.path.exists(data_file_folder):
    for file in os.listdir(data_file_folder):
        if not file.endswith(".xlsx"):
            continue
        print(f"loading {file}...")

        # search the file in the path, it need combine the path with file name
        file_path = os.path.join(data_file_folder, file)
        try:
            # sheet_name is the tab in execel
            df = pd.read_excel(file_path, sheet_name="Sale")
            all_dataframes.append(df)
        except ValueError:
            print(f"skipping {file} -- no 'Sale' Sheet")
    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        combined_df.to_excel("master file.xlsx", index=False)
        print(f"Successfully concatenated {len(all_dataframes)} files")
        print(combined_df.head())
    else:
        print(" No valid Excel files were loaded")
else:
    print(f"data folder does not exist")
