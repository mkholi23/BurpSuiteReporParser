import pandas as pd
import os
#Coded by ChatGPT 4, Assisted by Mark Kholi.
# Load the data
df = pd.read_excel("raw.xlsx", header=None)

# Initialize variables
current_asset = None
issue_name = None
last_issue_name = None
issue_severity = None
description_start = False
description = ''
issues_data = []
description_dict = {}

# Parse the data
for index, row in df.iterrows():
    if pd.isnull(row[0]):
        continue
    elif str(row[0]).startswith("Issues found on "):
        current_asset = str(row[0]).replace("Issues found on ", "").strip()
    elif "Issue description " in str(row[0]):
        description_start = True
    elif description_start and pd.notnull(df.loc[index+2, 0]):  # Adjusting index for description
        description_dict[last_issue_name] = str(df.loc[index+2, 0]).strip()  # Adjusting index for description
        description_start = False
    elif current_asset and " [1]" in str(row[0]):
        issue_name = str(row[0]).strip()
        last_issue_name = issue_name.replace(' [1]', '')  # Remove the trailing " [1]" here
        issue_severity = str(df.loc[index + 1, 1]).strip() if pd.notnull(df.loc[index + 1, 1]) else None
        description = description_dict.get(last_issue_name, "Not found")  # Use last_issue_name to lookup description
        issues_data.append([current_asset, last_issue_name, issue_severity, description])  # Use last_issue_name here

# Convert list to DataFrame
df_vulnerabilities = pd.DataFrame(issues_data, columns=["Asset", "Issue_Name", "Severity", "Description"])

# Write to a new excel file# Open the raw.xlsx file again
df_raw = pd.read_excel("raw.xlsx", header=None)

# Iterate through column A in raw.xlsx
for index, row in df_raw.iterrows():
    # Check if row[0] is an issue_name
    if row[0] in df_vulnerabilities['Issue_Name'].values:
        # Find the corresponding rows in df_vulnerabilities
        vuln_indexes = df_vulnerabilities[df_vulnerabilities['Issue_Name'] == row[0]].index.tolist()
        # Update all instances of the issue_name with the description from 5 cells below
        for vuln_index in vuln_indexes:
            df_vulnerabilities.loc[vuln_index, 'Description'] = df_raw.loc[index+5, 0]

# Write to a new excel file
severity_index = None
for index, row in df_raw.iterrows():
    if str(row[0]) == "Issues by severity":
        severity_index = index
        break

# Extract the severity information
severity_data = df_raw.iloc[severity_index:severity_index+6].copy()
severity_data.columns = severity_data.iloc[0]
severity_data = severity_data[1:]

# Write to a new excel file
with pd.ExcelWriter("output.xlsx") as writer:
    df_vulnerabilities.to_excel(writer, sheet_name="Vulnerabilities", index=False)
    severity_data.to_excel(writer, sheet_name="Dashboard", index=False)

# Open the file in Excel
os.system('open output.xlsx')
