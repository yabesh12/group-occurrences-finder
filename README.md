# Group Occurrences Finder

This script extracts and counts occurrences of groups from an Excel file based on a specified identifier pattern.

## Scope

The script is designed to process an Excel file with a specific column containing additional comments. It extracts groups identified by a regular expression pattern and counts their occurrences. The final output is saved in an Excel file.

## Instructions to Run

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yabesh12/group-occurrences-finder


2. **Create Virtualenv:**
   ```bash
   python -m venv venv


3. **Activate Virtuanenv:**
    ```bash
    source venv/bin/activate


4. **Install dependencies:**
    ```bash
    pip install -r requirements.txt


5. **Run the script:**
    ```bash
    python find_group_occurrences_by_specific_pattern.py



## Note
Retrieve the input file at runtime  
Default identifier pattern is - ```[code]<I>XXXX</I>[/code]```
