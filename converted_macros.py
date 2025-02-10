import pandas as pd
from datetime import datetime
import logging

class ConvertedExcelMacros:
    def __init__(self, dataframe):
        self.dataframe = dataframe
        logging.basicConfig(filename='converted_macros.log', level=logging.ERROR)

    def copy_data(self):
        try:
            if 'Data' in self.dataframe and 'Summary' in self.dataframe:
                self.dataframe['Summary'] = self.dataframe['Data'].copy()[:10]
                print("Data Copied Successfully!")
            else:
                raise KeyError("Required sheets 'Data' or 'Summary' are not available.")
        except Exception as e:
            logging.error(f"Error in copy_data function: {e}")

    def clear_summary(self):
        try:
            if 'Summary' in self.dataframe:
                self.dataframe['Summary'] = self.dataframe['Summary'].iloc[0:0]
                print("Summary Sheet Cleared!")
            else:
                raise KeyError("Sheet 'Summary' is not available.")
        except Exception as e:
            logging.error(f"Error in clear_summary function: {e}")

    def add_new_row(self):
        try:
            if 'Data' in self.dataframe:
                last_row = len(self.dataframe['Data'])
                new_data = {"Column1": "New Entry", "Column2": datetime.now()}
                self.dataframe['Data'] = self.dataframe['Data'].append(new_data, ignore_index=True)
                print("New Row Added!")
            else:
                raise KeyError("Sheet 'Data' is not available.")
        except Exception as e:
            logging.error(f"Error in add_new_row function: {e}")

    def operation(self):
        try:
            operations = {
                1: self.copy_data,
                2: self.clear_summary,
                3: self.add_new_row
            }
            print("\nSelect operation:")
            print("1. Copy Data")
            print("2. Clear Summary")
            print("3. Add New Row")
            choice = int(input("Enter your choice (1/2/3): ").strip())
            if choice in operations:
                operations[choice]()
            else:
                print("Invalid Choice!")
        except Exception as e:
            logging.error(f"Error in operation function: {e}")

def main():
    try:
        input_file = input("Enter the path of the Excel file: ").strip()
        df = pd.read_excel(input_file, sheet_name=None)  # Load all sheets into a dict of dataframes
        macro_operations = ConvertedExcelMacros(dataframe=df)
        macro_operations.operation()
        with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
            for sheet, data in df.items():
                data.to_excel(writer, sheet_name=sheet, index=False)  # Save updated data back to the same file
        print("Excel file updated successfully!")
    except Exception as e:
        logging.error(f"Error in main function: {e}")

if __name__ == "__main__":
    main()

