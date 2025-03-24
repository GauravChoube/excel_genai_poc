import pandas as pd
import threading
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Thread safety lock
lock = threading.Lock()

class ConvertedExcelMacros:

    def __init__(self, dataframes):
        self.dataframes = dataframes
        
        # Extract relevant DataFrames
        with lock:
            self.data_sheet = self.dataframes.get('Data', pd.DataFrame(columns=['Column1', 'Column2']))
            self.summary_sheet = self.dataframes.get('Summary', pd.DataFrame(columns=['Column1', 'Column2']))
            
            # Handle missing or empty DataFrames
            if self.data_sheet.empty:
                logging.warning("Data sheet is empty or missing. Initializing with default structure.")
            if self.summary_sheet.empty:
                logging.warning("Summary sheet is empty or missing. Initializing with default structure.")
    
    def existing_macros(self):
        return ['CopyData', 'ClearSummary', 'AddNewRow']

    def execute_macro(self, macro_name):
        try:
            if macro_name == 'CopyData':
                return self._copy_data()
            elif macro_name == 'ClearSummary':
                return self._clear_summary()
            elif macro_name == 'AddNewRow':
                return self._add_new_row()
            else:
                logging.error(f"Macro '{macro_name}' not found.")
                return f"Error: Macro '{macro_name}' not found."
        except Exception as e:
            logging.error(f"Error executing macro '{macro_name}': {str(e)}")
            return f"Error: {str(e)}"
    
    def _copy_data(self):
        with lock:
            try:
                if self.data_sheet.empty:
                    logging.warning("Data sheet is empty. Cannot copy data.")
                    return "Error: Data sheet is empty."
                self.summary_sheet = self.data_sheet.iloc[:10].copy()
                logging.info("Data copied from 'Data' sheet to 'Summary' sheet.")
                return "Data Copied Successfully!"
            except Exception as e:
                logging.error(f"Error during CopyData: {str(e)}")
                return f"Error: {str(e)}"
    
    def _clear_summary(self):
        with lock:
            try:
                if self.summary_sheet.empty:
                    logging.warning("Summary sheet is already empty.")
                    return "Summary sheet is already empty."
                self.summary_sheet = pd.DataFrame(columns=['Column1', 'Column2'])
                logging.info("Summary sheet cleared.")
                return "Summary Sheet Cleared Successfully!"
            except Exception as e:
                logging.error(f"Error during ClearSummary: {str(e)}")
                return f"Error: {str(e)}"
    
    def _add_new_row(self):
        with lock:
            try:
                last_index = len(self.data_sheet)
                new_row = {'Column1': 'New Entry', 'Column2': datetime.now()}
                self.data_sheet = pd.concat([self.data_sheet, pd.DataFrame([new_row])], ignore_index=True)
                logging.info("New row added to 'Data' sheet.")
                return "New Row Added Successfully!"
            except Exception as e:
                logging.error(f"Error during AddNewRow: {str(e)}")
                return f"Error: {str(e)}"


