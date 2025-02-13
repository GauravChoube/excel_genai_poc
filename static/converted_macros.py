from openpyxl import load_workbook
import pandas as pd
import logging

# Configure logging
logging.basicConfig(filename='macro_conversion.log', level=logging.ERROR, format='%(asctime)s - %(message)s')

class ConvertedExcelMacros:
    def __init__(self, dataframe, workbook_path):
        self.dataframe = dataframe
        self.workbook_path = workbook_path

    def copy_data(self):
        # Copies data from "Data" sheet to "Summary" sheet
        try:
            workbook = load_workbook(self.workbook_path, keep_vba=True)
            data_sheet = workbook["Data"]
            summary_sheet = workbook["Summary"]

            for row in data_sheet.iter_rows(min_row=1, max_row=10, min_col=1, max_col=2, values_only=True):
                summary_sheet.append(row)

            workbook.save(self.workbook_path)
            print("Data Copied Successfully!")
        except Exception as e:
            logging.error(f"Error in copy_data function: {str(e)}")
            print("Failed to copy data. See logs for details.")

    def clear_summary(self):
        # Clears data from "Summary" sheet
        try:
            workbook = load_workbook(self.workbook_path, keep_vba=True)
            summary_sheet = workbook["Summary"]

            for row in summary_sheet.iter_rows():
                for cell in row:
                    cell.value = None

            workbook.save(self.workbook_path)
            print("Summary Sheet Cleared!")
        except Exception as e:
            logging.error(f"Error in clear_summary function: {str(e)}")
            print("Failed to clear summary. See logs for details.")

    def add_new_row(self):
        # Adds a new row with a timestamp in the "Data" sheet
        try:
            workbook = load_workbook(self.workbook_path, keep_vba=True)
            data_sheet = workbook["Data"]

            last_row = data_sheet.max_row + 1
            data_sheet.cell(row=last_row, column=1).value = "New Entry"
            data_sheet.cell(row=last_row, column=2).value = pd.Timestamp.now()

            workbook.save(self.workbook_path)
            print("New Row Added!")
        except Exception as e:
            logging.error(f"Error in add_new_row function: {str(e)}")
            print("Failed to add new row. See logs for details.")

    def operation(self):
        # Ask user to select operation based on available Python functions
        while True:
            print("\nSelect Operation:")
            print("1: Copy Data")
            print("2: Clear Summary")
            print("3: Add New Row")
            print("4: Exit")
            choice = input("Enter your choice: ")

            if choice == "1":
                self.copy_data()
            elif choice == "2":
                self.clear_summary()
            elif choice == "3":
                self.add_new_row()
            elif choice == "4":
                print("Operation terminated by user.")
                break
            else:
                print("Invalid choice. Try again.")

if __name__ == "__main__":
    try:
        input_file = input("Enter the path to the Excel file (.xlsm): ")
        workbook_df = pd.read_excel(input_file, sheet_name=None)

        macro_converter = ConvertedExcelMacros(workbook_df, input_file)
        macro_converter.operation()

        # Update the same Excel file with the modified DataFrame while retaining macros
        writer = pd.ExcelWriter(input_file, engine="openpyxl", mode="a", if_sheet_exists="overlay")
        for sheet_name, df in workbook_df.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
        writer.book.save(input_file)

    except Exception as e:
        logging.error(f"Error in main function: {str(e)}")
        print("Failed to process the Excel file. See logs for details.")

