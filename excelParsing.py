import pandas as pd
import win32com.client
import os
import openpyxl
from openpyxl.utils import range_boundaries


class ExcelVBAProcessor:
    def __init__(self, file_path, openAIClient):
        self.file_path = file_path
        self.output_path = "converted_macros.py"
        self.excel_data = None
        self.vba_macros = None
        self.python_class_code = None
        self.openAIClient = openAIClient

    def read_excel_data(self):
        """Read all sheets from the Excel file into a dictionary."""
        self.excel_data = pd.read_excel(self.file_path, sheet_name=None)
        return self.excel_data
    
    def read_excel_with_formulas_all_sheets(self):
        """
        Reads all sheets of an Excel file, including formulas, into a dictionary of DataFrames.

        Args:
            file_path (str): Path to the Excel file.

        Returns:
            dict: A dictionary where keys are sheet names and values are DataFrames containing data and formulas.
        """
        try:
            # Load workbook without evaluating formulas
            wb = openpyxl.load_workbook(self.file_path, data_only=False)  # data_only=False to capture formulas
            sheet_names = wb.sheetnames
            all_dfs = {}

            # Iterate through each sheet
            for sheet_name in sheet_names:
                ws = wb[sheet_name]
                spaceCnt = 1
                # Dynamically calculate the used range
                range_string = ws.calculate_dimension()
                min_col, min_row, max_col, max_row = range_boundaries(range_string)

                data_rows = []
                header_row = None

                # Iterate over the entire range to capture values and formulas
                for row_idx, row in enumerate(ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col), start=1):
                    row_data = []
                    print([str(cell.value) for cell in row]) 
                    print([cell.data_type for cell in row])  # Check cell types
                    for cell in row:
                        print(f"{cell.col_idx}-{cell.internal_value}")
                        # Check if cell contains a formula
                        if cell.data_type == 'f':  # Formula cell
                            # row_data.append(f"={cell.value}")
                            row_data.append(str(cell.internal_value))
                        else:
                            if row_idx == 1 and cell.value == None:
                                cell.value = " "*spaceCnt
                                spaceCnt += 1
                            row_data.append(cell.value)

                    # Store header row separately
                    if row_idx == 1:
                        header_row = row_data
                    else:
                        data_rows.append(row_data)

                # Create DataFrame if valid data is present
                if header_row and data_rows:
                    df = pd.DataFrame(data_rows, columns=header_row)
                    print(df)
                    all_dfs[sheet_name] = df
                else:
                    all_dfs[sheet_name] = pd.DataFrame()  # Empty DataFrame for empty sheets

            return all_dfs

        except FileNotFoundError:
            print(f"Error: File not found at {self.file_path}")
            return None
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def extract_vba_macros(self):
        """Extract all VBA macros from the Excel file."""
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Keep Excel hidden

        try:
            workbook = excel.Workbooks.Open(self.file_path)
            vba_project = workbook.VBProject  # Access the VBA project
            self.vba_macros = {}

            for i in range(vba_project.VBComponents.Count):
                module = vba_project.VBComponents.Item(i + 1)
                if module.Type == 1:  # Standard module
                    module_name = module.Name
                    code_lines = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
                    self.vba_macros[module_name] = {
                        "code": code_lines,
                        "macros": self.extract_macro_names(code_lines)  # Extract macro names
                    }

            workbook.Close(SaveChanges=False)
            return self.vba_macros

        except Exception as e:
            print(f"Error extracting VBA macros: {e}")
            return None

        finally:
            excel.Quit()

    def extract_macro_names(self, code):
        """Extract macro names from the VBA code."""
        import re
        macro_names = re.findall(r'(?<=Sub\s)(\w+)', code)  # Find all macro names
        return macro_names

    def convert_vba_to_python(self):
        """Convert VBA macros to Python class methods."""
        if not self.vba_macros:
            print("No VBA macros found.")
            return None
        # macrosCodeStr = "```\n"
        # for module, code in self.vba_macros.items():
        #     macrosCodeStr += code +"\n\n"
        # macrosCodeStr +="```"
        # print(f"macrosCodeStr:{macrosCodeStr}")
        class_code = ""
        # class_code += "    def __init__(self, file_path):\n"
        # class_code += "        self.file_path = file_path\n\n"

        for module, code in self.vba_macros.items():
            class_code += self.vba_to_python_translator(code) + "\n\n"

        # class_code += "    def process_file(self):\n"
        # class_code += "        print(\"Select operation to perform:\")\n"
        # class_code += "        while True:\n"
        # class_code += "            print(\"Available Functions:\")\n"
        # class_code += "            methods = [method for method in dir(self) if callable(getattr(self, method)) and not method.startswith('__')]\n"
        # class_code += "            for i, method in enumerate(methods):\n"
        # class_code += "                print(f'{i + 1}. {method}')\n"
        # class_code += "            choice = input(\"Enter function number to execute (or 'exit' to quit): \")\n"
        # class_code += "            if choice.lower() == 'exit':\n"
        # class_code += "                break\n"
        # class_code += "            if choice.isdigit() and 1 <= int(choice) <= len(methods):\n"
        # class_code += "                getattr(self, methods[int(choice) - 1])()\n"
        # class_code += "            else:\n"
        # class_code += "                print(\"Invalid choice, please try again.\")\n\n"

        # class_code += "if __name__ == \"__main__\":\n"
        # class_code += "    file_path = input(\"Enter the Excel file path: \")\n"
        # class_code += "    macros_instance = ConvertedExcelMacros(file_path)\n"
        # class_code += "    macros_instance.process_file()\n"

        self.python_class_code = class_code
        return self.python_class_code

    def vba_to_python_translator(self, vba_code):
        """Convert VBA to Python using OpenAI API."""

        print(f"VBA code:\n{vba_code}")

        # read a promt from file
        fp = open("./prompt.txt", "r")
        prompt = fp.read()
        prompt = prompt + f"\n'''\n{vba_code}\n'''"

        print(f"Final prompt as follow:=>\n{prompt}")

        python_code = self.openAIClient.promptCall(prompt)

        print(f"Converted Python code:\n{python_code}")
        print(f"========================================")
        # return "    " + python_code.replace("\n", "\n    ")  # Indent properly
        return python_code

    def formula_to_python_translator(self, formulaStr):
        """Convert formulaStr to Python using OpenAI API."""

        print(f"VBA code:\n{formulaStr}")

        # read a promt from file
        fp = open("./prompt_formula.txt", "r")
        prompt = fp.read()
        prompt = prompt + f"\n'''\n{formulaStr}\n'''"

        print(f"Final prompt as follow:=>\n{prompt}")

        python_code = self.openAIClient.promptCall(prompt)

        print(f"Converted Python code:\n{python_code}")
        print(f"========================================")
        # return "    " + python_code.replace("\n", "\n    ")  # Indent properly
        return python_code

    def save_python_class(self):
        """Save the generated Python class to a .py file."""
        if not self.python_class_code:
            print("No Python class generated to save.")
            return
        with open(self.output_path, "w", encoding="utf-8") as file:
            file.write(self.python_class_code)
        print(f"✅ Python class saved successfully: {self.output_path}")

    def process_excel_file(self):
        """ Execute the entire workflow: Read data, extract VBA, convert to Python, consolidate macros. """
        print("Reading Excel data...")
        self.read_excel_data()

        print("Extracting VBA macros...")
        self.extract_vba_macros()

        print("Converting VBA macros to Python...")
        self.convert_vba_to_python()

        print("💾 Saving Python class to file...")
        self.save_python_class()


def main():
    """Main function to process the Excel file and save the converted Python class."""
    file_path = r"C:\Users\gaurav.j.choubey\Desktop\project\gtic-7-2025\excelsheet_poc\excel_genai_poc\examplemacro1.xlsm"
    openAIClient = None  # Replace with actual OpenAI client instance

    processor = ExcelVBAProcessor(file_path, openAIClient)
    processor.process_excel_file()

    print("Running converted Python class...")
    os.system(f"python {processor.output_path}")


if __name__ == "__main__":
    main()
