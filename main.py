import os
import pandas as pd


def format_size(size):
    # Define the threshold for KB, MB, GB
    if size < 1024:
        return f"{size} Bytes"
    elif size < 1024**2:
        return f"{size / 1024:.2f} KB"
    elif size < 1024**3:
        return f"{size / 1024**2:.2f} MB"
    else:
        return f"{size / 1024**3:.2f} GB"


def get_files_info(directory):
    # Create an empty list for file information
    files_info = []

    # Walk through the directory
    for root, dirs, files in os.walk(directory):
        for file in files:
            # Get the full path of the file
            full_path = os.path.join(root, file)
            # Get the size of the file
            size = os.path.getsize(full_path)
            # Format the size
            formatted_size = format_size(size)
            # Append the file name and size to the list
            files_info.append((file, formatted_size))

    # Create a DataFrame from the list
    df = pd.DataFrame(files_info, columns=["File Name", "Size"])
    return df


def main():
    # Prompt the user for the directory path
    directory = input("Enter the directory path: ")

    # Get the DataFrame of file info
    df = get_files_info(directory)

    # Prompt for output file name
    output_filename = input("Enter the output file name: ")

    # Check if the filename ends with '.xlsx' and append if not
    if not output_filename.lower().endswith(".xlsx"):
        output_filename += ".xlsx"

    # Export the DataFrame to an Excel file
    df.to_excel(output_filename, index=False)
    print(f"Data exported to {output_filename}")


if __name__ == "__main__":
    main()
