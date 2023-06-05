
# Burp Suite Report Parser (BSparser)

This Python script parses Burp Suite reports, extracting key information and saving it in a more structured format.

## Overview

This script reads data from an excel file (raw.xlsx), extracts various issue details like Asset, Issue Name, Severity, and Description, and writes the processed data into a new excel file (output.xlsx).

## How It Works

The script initializes a few variables and then iterates over the rows of the input excel file. 

It identifies the start of a new asset by the marker "Issues found on", extracts issue names, and keeps track of their descriptions. It also records the severity level of each issue. 

Next, it creates a pandas DataFrame from the collected data and then opens the input file again to update the issue descriptions. 

Finally, it extracts severity information from the raw data and writes both the vulnerability details and severity information into the output excel file in two separate sheets: "Vulnerabilities" and "Dashboard".

## Usage

1. Make sure that pandas is installed in your Python environment. If not, install it using pip:

   ```
   pip install pandas
   ```

2. Place the `raw.xlsx` file (your Burp Suite report) in the same directory as the script.

3. Run the script:

   ```
   python script_name.py
   ```

Replace `script_name.py` with the actual name of your script. 

4. The output will be saved in the same directory as `output.xlsx`.

## Credits

This script was coded by OpenAI's ChatGPT-4, with assistance from Mark Kholi.

## Disclaimer

This script assumes a specific structure of the Burp Suite report. If your report is structured differently, you might need to adjust the script accordingly.

---

Remember, `README.md` files on GitHub support Markdown syntax, which allows you to add formatting like headers, links, lists, etc. to make your README files easier to read.
