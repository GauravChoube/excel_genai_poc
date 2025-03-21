# logging setup
import logging
from threading import Lock

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

import pandas as pd
from datetime import datetime

class ConvertedExcelMacros:
    def __init__(self, df_dict):
        self.df_dict = df_dict
        self.lock = Lock()

        # Check and initialize required DataFrames
        self._initialize_dataframe("Data", ["Column1", "Column2"])
        self._initialize_dataframe("Summary", ["Column1", "Column2"])

    def _initialize_dataframe(self, key, columns):
        """
        Initializes a DataFrame if it is missing or empty.
        """
        if key not in self.df_dict or self.df_dict[key].empty:
            logging.warning(f"{key} DataFrame is missing or empty. Initializing with default structure.")
            self.df_dict[key] = pd.DataFrame(columns=columns)

    def existing_macros(self):
        """
        Returns a list of existing macro names.
        """
        return ["CopyData", "ClearSummary", "AddNewRow"]

    def execute_macro(self, macro_name):
        """
        Executes the specified macro on the relevant DataFrame.

        Args:
            macro_name: Name of the macro to execute.

        Returns:
            Result of the macro execution, if applicable.
        """
        macro_map = {
            "CopyData": self._copy_data,
            "ClearSummary": self._clear_summary,
            "AddNewRow": self._add_new_row
        }

        if macro_name not in macro_map:
            logging.error(f"Macro '{macro_name}' does not exist.")
            return f"Error: Macro '{macro_name}' does not exist."

        try:
            with self.lock:
                return macro_map[macro_name]()
        except Exception as e:
            logging.error(f"Error during macro execution '{macro_name}': {e}")
            return f"Error: {e}"

    def _copy_data(self):
        """
        Copies data from 'Data' to 'Summary' and returns a success message.
        """
        if self.df_dict["Data"].empty:
            logging.warning("Data DataFrame is empty. Cannot perform CopyData macro.")
            return "Warning: Data DataFrame is empty. Cannot perform CopyData."

        self.df_dict["Summary"] = self.df_dict["Data"].copy()
        logging.info("Data copied successfully to Summary.")
        return "Data Copied Successfully!"

    def _clear_summary(self):
        """
        Clears data from 'Summary' DataFrame and returns a success message.
        """
        self.df_dict["Summary"] = pd.DataFrame(columns=self.df_dict["Summary"].columns)
        logging.info("Summary DataFrame cleared successfully.")
        return "Summary Sheet Cleared!"

    def _add_new_row(self):
        """
        Adds a new row with a timestamp to 'Data' and returns a success message.
        """
        new_row = {"Column1": "New Entry", "Column2": datetime.now()}
        self.df_dict["Data"] = pd.concat([self.df_dict["Data"], pd.DataFrame([new_row])], ignore_index=True)
        logging.info("New row added to Data DataFrame successfully.")
        return "New Row Added!"





