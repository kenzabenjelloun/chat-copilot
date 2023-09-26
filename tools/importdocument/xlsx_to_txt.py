import openpyxl
import os
import subprocess


def create_text_file(data, output_folder):
    file_paths = []
    for i, row in enumerate(data, start=1):
        txt_file_name = os.path.join(output_folder, f'BaseQnASecurity_{i}.txt')
        with open(txt_file_name, 'w') as txt_file:
            txt_file.write('\n'.join(map(str, row)))
        file_paths.append(txt_file_name)
    return 

def main(input_excel, output_folder, sheet_name):
    # Load the Excel file
    workbook = openpyxl.load_workbook(input_excel)
    sheet = workbook[sheet_name]

    data = []

    # Iterate through rows and cells to retrieve data
    for row in sheet.iter_rows():
        row_data = [cell.value for cell in row]
        data.append(row_data)

    create_text_file(data, output_folder)

if __name__ == "__main__":
    input_excel = "C:\\Users\\t-kenzab\\OneDrive - Microsoft\\Bureau\\Generative AI\\BaseQnASecurity.xlsx"  
    output_folder = "C:\\Users\\t-kenzab\\OneDrive - Microsoft\\Bureau\\Generative AI\\secu_folder"  # Output folder where text files will be saved
    sheet_name = "CyberVadis_Data"  # Name of the sheet you want to use

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    main(input_excel, output_folder, sheet_name)

    for i in range(1, len(os.listdir(output_folder)) + 1):
        command = f'dotnet run --files "{os.path.join(output_folder, f"BaseQnASecurity_{i}.txt")}"'
        subprocess.run(command, shell=True)