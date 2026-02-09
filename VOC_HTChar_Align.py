"""
A script used for inter-sample data alignment and classification.

Author: Andriy Rebryk, andrew.rebryk@gmail.com

To run the script, Python 3 (https://www.python.org/downloads/windows/),
all the necessary modules (see lines 37-50),
and Google Chrome have to be installed.

Modules can be intalled via Command Prompt as follows:
Type “cmd” in the Search line --> Click on the Command Prompt icon -->
--> Type: pip install module_name --> Press “Enter” --> Repeat for the next module.
Module names:
tk
pandas
openpyxl
pubchempy
selenium

Before running the script, change the path to the working folder and the output file name (see line 434).

The script should be executed in Command Prompt.
For this, the following steps should be applied:

1. Type “cmd” in the Search line --> Click on the Command Prompt icon -->
--> Type: python "the path to the script including the extension (.py)", e.g. python "D:/Projects/VOC_HTChar_Align.py" -->
--> Press “Enter”.

2. Navigate to the folder where the file to be processed is stored, select the file, and press “Open”.
The processed file will be saved to the location you specified (line 434).

This script is tailored to a certain data table format.
Feel free to change the script or use any part of it for your needs.
"""

#importing modules
import tkinter as tk
import tkinter.filedialog
from tkinter.filedialog import askopenfilename
import time
import pandas as pd
from pandas import DataFrame
import numpy as np
import pubchempy as pcp
from pubchempy import get_compounds, Compound
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

root = tk.Tk()
root.withdraw() #prevents the tkinter window to come up
#opening the file
xls_path = askopenfilename(title = 'Select Excel file for processing...')
root.destroy() #closing the window

#adding processing start time
startTime = time.strftime("%Y-%m-%d %H:%M:%S")
print(' ')
print('Processing start time:', startTime)
print(' ')

#storing the file as DataFrame (DF)
inXSL = pd.read_excel(xls_path, header = None, sheet_name = None)

#getting Excel sheet (sample) names; will be needed to add suffixes to column names for each sample
sheets = pd.ExcelFile(xls_path).sheet_names


#creating empty lists/arrays for further data appending
names = []
sheet_names = []
df_list = list()

#creating a list with columns to which sample name suffixes should NOT be added
#these are the columns that will be used as base for joining (merging) the samples
keep_same = {'Substance', 'CAS', 'RT_int'}

#creating separate DataFrames (DFs) for each sample (Excel sheet)
for i, sheet in enumerate(sheets):
    #storing each sheet separately as DF; DF name indexing starts from 1 not 0, i.e. from df1 to df36 in this case
    globals()['df' + str(i + 1)] = pd.read_excel(xls_path, header = None, sheet_name = sheet)
    #deleting column '#' & setting the 1st row from the initial sheet ('Substance', 'CAS', 'Similarity, %', 'RT, min', 'Area') as a header
    globals()['df' + str(i + 1)] = globals()['df' + str(i + 1)].drop(globals()['df' + str(i + 1)].columns[[0]], axis = 1).rename(columns = globals()['df' + str(i + 1)].iloc[0]).iloc[1:].reset_index(drop = True)

    #QC printing
    print('Input sheet', str(sheets[i]), '\n', globals()['df' + str(i + 1)])
    print(150*'*')
    
    names.append(str('df' + str(i + 1))) #appending DF names to an empty list 'names'
    sheet_names.append(str(sheets[i])) #appending Excel sheet names to an empty list 'sheet_names'
    #adding sample name suffixes to all columns except for 'Substance', 'CAS' & 'RT_int'
    globals()['df' + str(i + 1)].columns = ['{}{}'.format(c, '' if c in keep_same else str('_' + str(sheets[i]))) for c in globals()['df' + str(i + 1)].columns]
    #appending all DFs to an empty DF list 'df_list'
    df_list.append(globals()['df' + str(i + 1)])

#QC printing
print('Input sheet names:\n', sheet_names)
print(150*'*')
print('All-DF list before processing:\n', df_list)
print(150*'*')


def correct_cas_number(cas):
    if not isinstance(cas, str):
        cas = str(cas)  #converting CAS# to string
    if cas == "0-0-0": #replacing '0-0-0' with 'NA'
        return "NA"
    parts = cas.split("-")
    if len(parts) == 3:
        #checking for valid CAS structure: first & last parts must be digits
        if parts[0].isdigit() and parts[2].isdigit():
            if len(parts[1]) == 1:  #correcting cases like "110-0-9" to "110-00-9"
                parts[1] = f"0{parts[1]}"
            #returning the corrected CAS
            return "-".join(parts)
    #if no correction needed, returning the CAS as is
    return cas

#correcting CAS numbers in a list of DFs
def correct_cas_in_df_list(df_list):
    for i, df in enumerate(df_list):
        if "CAS" in df.columns:
            df["CAS"] = df["CAS"].apply(correct_cas_number)
        else:
            print(f"Warning: DataFrame {i+1} does not contain a 'CAS' column.\n")
    return df_list

#correct CAS in all DFs in df_list
df_list = correct_cas_in_df_list(df_list)
#QC printing
print('DF list w/ corrected CAS:\n', df_list)
print(150*'*')


#aligning multiple DFs based on approximate matching for 'RT, min' & exact matching for 'Substance' & 'CAS' columns
def align_DFs(dataframes, tolerance = 0.1):
    #dataframes is a dictionary where keys are DF names & values are DFs
    #tolerance is a value for matching 'RT, min'; 0.1 min is a default value

    #creating an empty list for aligned data
    aligned_data = []

    #getting the list of unique "Substance" & "CAS" pairs across all DFs
    unique_keys = set()
    for df_name, df in dataframes.items():
        for _, row in df.iterrows():
            unique_keys.add((row["Substance"], row["CAS"]))

    #iterating over each unique "Substance" & "CAS" pair
    for substance, cas in unique_keys:
        row_base = {"Substance": substance, "CAS": cas}
        rt_values = []

        #collecting RT values for each DF
        for df_name, df in dataframes.items():
            rt_col = next((col for col in df.columns if col.startswith("RT, min_")), None)
            sim_col = next((col for col in df.columns if col.startswith("Similarity, %_")), None)
            area_col = next((col for col in df.columns if col.startswith("Area_")), None)

            #filtering rows matching the current "Substance" & "CAS"
            match = df[(df["Substance"] == substance) & (df["CAS"] == cas)]
            
            if not match.empty:
                for _, match_row in match.iterrows():
                    if rt_col:
                        rt_values.append((df_name, match_row[rt_col], match_row[sim_col] if sim_col else None, match_row[area_col] if area_col else None))

        #sorting the RT values by "RT, min" value (second element)
        rt_values.sort(key = lambda x: x[1])

        #creating lists to keep track of groups
        grouped_rows = []
        current_group = [rt_values[0]]  #starting the first group with the minimum "RT, min" value
        min_rt_value = rt_values[0][1]

        #grouping RT values within tolerance range
        for i in range(1, len(rt_values)):
            df_name, rt_value, sim_value, area_value = rt_values[i]
            if abs(rt_value - min_rt_value) <= tolerance:
                current_group.append((df_name, rt_value, sim_value, area_value))
            else:
                #creating a row for the previous group
                grouped_rows.append(current_group)
                current_group = [(df_name, rt_value, sim_value, area_value)]  #starting a new group
                min_rt_value = rt_value

        #adding the last group
        grouped_rows.append(current_group)

        #creatinge a separate entry for each group
        for group in grouped_rows:
            new_row = row_base.copy()
            for df_name, rt_value, sim_value, area_value in group:
                new_row[f"RT, min_{df_name}"] = rt_value
                new_row[f"Similarity, %_{df_name}"] = sim_value
                new_row[f"Area_{df_name}"] = area_value
            aligned_data.append(new_row)

    #creating a DF from the aligned data
    df_result = pd.DataFrame(aligned_data)
    return df_result

#generating a dictionary with names based on indices or custom names
dataframes = {} #creating an empty dictionary to hold DFs with suffix-based keys

#iterating over the df_list & corresponding sheet names
for i, (df, sheet) in enumerate(zip(df_list, sheets)):
    #using the sheet name as the key for the dictionary
    dataframes[sheet] = df

print('DF dictionary before alignment:\n', dataframes)
print(150*'*')

#running the alignment function 'align_DFs' with specified RT tolerance
df_aligned = align_DFs(dataframes, tolerance = 0.1667) #10 sec = ca. 0.1667 min
#QC printing
print('Aligned DF:\n', df_aligned)
print(150*'*')


#defyining the function to convert 'Substance' column values to InChIKey values
#InChIKeys are input for further batch classification at:
#https://cfb.fiehnlab.ucdavis.edu/
def Name_InChIKey_PubChem(IDs):
    try:
        results = pcp.get_compounds(IDs, 'name')
        for compound in results:
            return compound.inchikey
    except:
        return ''

#creating empty InChIKey column
df_aligned['InChIKey'] = pd.Series(dtype = 'float').fillna('')
#QC printing
print('Aligned DF w/ empty InChIKey column:\n', df_aligned)
print(150*'*')

indx = dict() #creating an empty dictionary to store DF indices for iteration

#searching for 'Substance' column entries using PubChem webpage
#and writing retrieved InChIKeys to an empty column
print('Retrieving InChIKeys from PubChem. It will take a while...\n')

for id in df_aligned['Substance']:
    indx[id] = df_aligned[df_aligned['Substance'] == id].index
    print(indx[id], '', id, '-->', Name_InChIKey_PubChem(id))
    df_aligned.loc[indx[id], 'InChIKey'] = Name_InChIKey_PubChem(id)

#QC printing
print(150*'*')
print('Aligned DF w/ InChIKeys:\n', df_aligned)
print(150*'*')

#re-sorting table by 'InChIKey' column in descending fashion, N/As are moved to the bottom of the table
#this step is needed for further classificetion script to skip N/As
df_aligned = df_aligned.sort_values(by = 'InChIKey', ascending = False, na_position = 'last')
df_aligned = df_aligned.reset_index(drop = True)
#QC printing
print('Sorted aligned DF by InChIKey:\n', df_aligned)
print(150*'*')

#creating a list of InChIKeys
inchikeys = df_aligned['InChIKey']

print('Retrieving class data from ClassyFire. It will take a while...\n')

#launching Chrome browser
driver = webdriver.Chrome()
#navigating to the website
driver.get('https://cfb.fiehnlab.ucdavis.edu/')

try:
    #waiting for the search text area and the ClassyFy button;
    #the time can be adjusted based on page loading time
    search_textarea = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'inchikeys')))
    classify_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']")))

    #inserting InChIKeys from the list
    search_textarea.clear() #clearing the input field
    for key in df_aligned['InChIKey'].dropna().tolist(): #.dropna().tolist() is used to skip N/As, otherwise, the script crashes
        search_textarea.send_keys(key + "\n")

    #clicking the ClassyFy button after entering all InChIKeys
    classify_button.click()

    #waiting for the results to load; the time can be adjusted based on page loading time,
    #22000 sec are used to be sure the processing will be done before the script exits the webpage
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'results')))
    table = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, 'table')))
    reset_button = WebDriverWait(driver, 22000).until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-warning')))

    #extracting data from the result table
    table_rows = table.find_elements(By.TAG_NAME, "tr") #finding all rows
    data = [] #creating an empty array to store the data
    for row in table_rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        cols = [col.text for col in cols]
        data.append(cols)
    print("*" * 150)

    #creating a list of column names; max. 11 columns are returned by ClassyFire
    column_names = ['InChIKey-2', 'Status', 'Kingdom', 'Superclass', 'Class', 'Subclass', 
                    'Parent Level 1', 'Parent Level 2', 'Parent Level 3', 'Parent Level 4', 'Parent Level 5']
    #adjusting the data to align with the specified column names
    filled_data = [row + [None] * (len(column_names) - len(row)) for row in data]
    #converting the data to DF & adding specified column names
    df_classyfy = pd.DataFrame(filled_data, columns = column_names)
    #dropping the row where all values are None & resetting indices
    df_classyfy.dropna(how = 'all', inplace = True)
    df_classyfy.reset_index(drop = True, inplace = True)
    #QC printing
    print('ClassiFyed table:\n', df_classyfy)
    print("*" * 150)

finally:
    #closing the browser
    driver.quit()

#concatenating the ClassyFyed table and the initial table
df_out = pd.concat([df_aligned, df_classyfy], axis = 1)
#dropping duplicate / unnecessary columns
df_out = df_out.drop(['InChIKey-2', 'Status'], axis = 1)
#copying data to output table & replacing possible N/A / NaN / None values with empty cells
df_out = df_out.fillna('')
#resetting indices
df_out = df_out.reset_index(drop = True)
#QC printing
print('Aligned & classified DF:\n', df_out)
print(150*'*')

#getting range and mean values for 'RT, min' & 'Similarity, %'
#finding all 'RT, min' & 'Similarity, %' columns
similarity_columns = [col for col in df_out.columns if col.startswith('Similarity, %')]
rt_columns = [col for col in df_out.columns if col.startswith('RT, min')]

#making all values in 'RT, min' & 'Similarity, %' numeric
df_out[similarity_columns] = df_out[similarity_columns].apply(pd.to_numeric, errors = 'coerce')
df_out[rt_columns] = df_out[rt_columns].apply(pd.to_numeric, errors = 'coerce')

#calculating range & mean values for 'Similarity, %'
df_out['Similarity_range, %'] = df_out[similarity_columns].apply(lambda row: f"{int(row.min())} - {int(row.max())}", axis = 1)
df_out['Similarity_mean, %'] = df_out[similarity_columns].mean(axis = 1).round(1)
#QC printing
print('DF w/ range & mean values for Similarity:\n', df_out)
print(150*'*')

#calculating range and mean values for 'RT, min'
df_out['RT_range, min'] = df_out[rt_columns].apply(lambda row: f"{row.min():.3f} - {row.max():.3f}", axis = 1)
df_out['RT_mean, min'] = df_out[rt_columns].mean(axis = 1).round(3)
#QC printing
print('DF w/ range & mean values for RT:\n', df_out)
print(150*'*')

#getting number of detections & detection frequency
#finding all columns that contain the word "Area" in the header & counting them
area_columns = [col for col in df_out.columns if "Area" in col]
area_total_count = len(area_columns)
print('Area column count:\n', area_total_count)
print(150*'*')

#checking for valid numeric or text values in "Area" columns & count "Number of detections"
def count_valid_detections(row):
    count = 0
    for col in area_columns:
        value = row[col]
        if isinstance(value, (int, float)) and not np.isnan(value):  #checking for valid numeric values
            count += 1
        elif isinstance(value, str) and value.strip():  #checking for non-empty strings
            count += 1
    return count

df_out["Number of detections"] = df_out.apply(count_valid_detections, axis = 1)

#calculating detection frequency ("DF, %") as a percentage of detections to total count "Area" columns
df_out["DF, %"] = ((df_out["Number of detections"] / area_total_count) * 100).round(0).astype(int)
#displaying the result
print('DF w/ detection frequency:\n', df_out)
print(150*'*')

#re-ordering columns
#columns to move and their new positions
columns_to_move = [
    'Number of detections',
    'DF, %',
    'RT_mean, min',
    'RT_range, min',
    'Similarity_mean, %',
    'Similarity_range, %'
]
new_pos = 2  #new position for the 1st column in the list; counting starts from 0

#getting all current columns as a list
columns = df_out.columns.tolist()

#removing the columns to move from their current positions
for col in columns_to_move:
    columns.remove(col)

#inserting the columns at the desired position
for i, col in enumerate(columns_to_move):
    columns.insert(new_pos + i, col)

#re-ordering the DataFrame
df_out = df_out[columns]

#moving 'InChIKey' column
#getting all current columns as a list
columns = df_out.columns.tolist()

#removing 'InChIKey' from the current position
columns.remove('InChIKey')

#inserting 'InChIKey' at the desired position (2nd; counting starts from 0)
columns.insert(4, 'InChIKey')

#re-ordering the DataFrame
df_out = df_out[columns]

#re-sorting the output table by 'RT_int, min' & 'Substance' columns  in ascending fashion
df_out = df_out.sort_values(by = ['RT_mean, min', 'Substance'], ascending = True)
#resetting indices
df_out = df_out.reset_index(drop = True)
#QC printing
print('Out file:\n', df_out)
print(150*'*')

#writing resulting DF to an Excel file
"""
Change the path according to your working folder path & output file name.
Do not change the file extension, i.e. .xlsx.
"""
df_out.to_excel('D:/Projects/Aligned_&_ClassiFyed.xlsx', sheet_name = 'OUT', index = False)

print('Processing start time:', startTime)
print('Processing end time:', time.strftime("%Y-%m-%d %H:%M:%S"))
print(150*'*')
print(150*'*')