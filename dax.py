import json
import pandas as pd

# Load the JSON data from a file
with open('th.json', 'r') as file:
    data = json.load(file)

# Initialize a dictionary to hold measures by display folder
measures_by_folder = {}

# Extract measures from the JSON structure
measures = data['createOrReplace']['database']['model']['tables'][0]['measures']

# Populate the dictionary based on the displayFolder
for measure in measures:
    name = measure["name"]
    expression = measure["expression"]
    display_folder = measure.get("displayFolder", "Common")  # Default to "Common" if not specified

    # Split the display folder to get the root folder
    root_folder = display_folder.split("\\")[0]  # Get the first part before any backslashes

    # Initialize the dictionary for this folder if it doesn't exist
    if root_folder not in measures_by_folder:
        measures_by_folder[root_folder] = {
            "NUM_DAX": [],
            "NUM_EXPRESSION": [],
            "DENOM_DAX": [],
            "DENOM_EXPRESSION": [],
            "BASELINE_DAX": [],
            "BASELINE_EXPRESSION": [],
            "TARGET_DAX": [],
            "TARGET_EXPRESSION": [],
            "INDICATOR_DAX": [],
            "INDICATOR_EXPRESSION": []
        }

    # Categorize the measure based on its name and expression
    if name.startswith("NUM"):
        measures_by_folder[root_folder]["NUM_DAX"].append(name)
        measures_by_folder[root_folder]["NUM_EXPRESSION"].append(expression)
    elif name.startswith("DENOM"):
        measures_by_folder[root_folder]["DENOM_DAX"].append(name)
        measures_by_folder[root_folder]["DENOM_EXPRESSION"].append(expression)
    elif name.startswith("WBP-B"):
        measures_by_folder[root_folder]["BASELINE_DAX"].append(name)
        measures_by_folder[root_folder]["BASELINE_EXPRESSION"].append(expression)
    elif name.startswith("WBP-T"):
        measures_by_folder[root_folder]["TARGET_DAX"].append(name)
        measures_by_folder[root_folder]["TARGET_EXPRESSION"].append(expression)
    else:
        measures_by_folder[root_folder]["INDICATOR_DAX"].append(name)
        measures_by_folder[root_folder]["INDICATOR_EXPRESSION"].append(expression)

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter("measures_separated_dax.xlsx", engine='xlsxwriter') as writer:
    # Write each folder's measures to a separate sheet
    for folder, measures in measures_by_folder.items():
        # Create a DataFrame for the measures
        df = pd.DataFrame({
            "NUM_DAX": pd.Series(measures["NUM_DAX"]),
            "NUM_EXPRESSION": pd.Series(measures["NUM_EXPRESSION"]),
            "DENOM_DAX": pd.Series(measures["DENOM_DAX"]),
            "DENOM_EXPRESSION": pd.Series(measures["DENOM_EXPRESSION"]),
            "BASELINE_DAX": pd.Series(measures["BASELINE_DAX"]),
            "BASELINE_EXPRESSION": pd.Series(measures["BASELINE_EXPRESSION"]),
            "TARGET_DAX": pd.Series(measures["TARGET_DAX"]),
            "TARGET_EXPRESSION": pd.Series(measures["TARGET_EXPRESSION"]),
            "INDICATOR_DAX": pd.Series(measures["INDICATOR_DAX"]),
            "INDICATOR_EXPRESSION": pd.Series(measures["INDICATOR_EXPRESSION"])
        })
        
        # Write to a sheet named after the folder
        df.to_excel(writer, sheet_name=folder, index=False)

print("Excel file 'measures_separated.xlsx' created successfully with separate sheets.")