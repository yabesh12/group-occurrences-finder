import pandas as pd
import re

def get_group_occurrences(file_path, identifier_pattern):
    """
    Read an Excel file and extract group occurrences based on the specified identifier pattern.

    Parameters:
    - file_path (str): The path to the Excel file.
    - identifier_pattern (str) (hardcoded): The regular expression pattern for identifying groups.

    Returns:
    - dict: A dictionary containing unique group occurrences.
    """
    try:
        df = pd.read_excel(file_path)

        # Convert the "Additional comments" column to strings
        df['Additional comments'] = df['Additional comments'].astype(str)

        # Extract groups using the specified identifier pattern
        groups_series = df['Additional comments'].apply(lambda x: re.findall(identifier_pattern, x))

        # Combine multiple groups in a single row and split by commas
        groups_list = groups_series.explode().str.strip()

        # Count the occurrences of each group
        group_occurrences = groups_list.value_counts().to_dict()

        # Merge the duplicate group occurrences
        unique_group_occurrences = {}
        for key, value in group_occurrences.items():
            keys = [k.strip() for k in key.split(',')]
            for k in keys:
                unique_group_occurrences[k] = unique_group_occurrences.get(k, 0) + value

    except Exception as e:
        print(f"Error Occured: {e}")
        unique_group_occurrences = {}

    return unique_group_occurrences

def main():
    """
    Main function to prompt user input, process data, and save results to an Excel file.
    """
    file_path = input("Enter the file path: ")
    identifier_pattern = r'\[code\]<I>(.*?)<\/I>\[\/code\]'

    result = get_group_occurrences(file_path, identifier_pattern)

    # Create a DataFrame from the result
    df_output = pd.DataFrame(list(result.items()), columns=['Group_name', 'Number of occurrences'])

    # Save the DataFrame to an Excel file
    excel_output_path = "output.xlsx"
    df_output.to_excel(excel_output_path, index=False)

    print(f"Output written to {excel_output_path}")
    
    return None

if __name__ == "__main__":
    main()
