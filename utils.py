import pandas as pd


class AllContracts:
    """
    A class to handle contract data processing from an Excel file using pandas.
    """

    def __init__(self, file_path):
        """
        Initializes the AllContracts object and loads the Excel data.

        Args:
            file_path (str): The path to the Excel file.

        Raises:
            FileNotFoundError: If the file cannot be found at the specified path.
            Exception: For other errors during file loading or processing.
        """
        self.file_path = file_path
        self.df = None
        try:
            # Use pandas to read the excel file. 'openpyxl' is needed for .xlsx files.
            self.df = pd.read_excel(self.file_path, engine="openpyxl")
        except FileNotFoundError:
            raise FileNotFoundError(f"Error: File not found at {self.file_path}")
        except Exception as e:
            raise Exception(f"An error occurred while loading the data: {e}")

    def payment_plan_name_counts(self):
        """
        Filters for 'Zone' payment plans and returns their value counts using pandas.

        Returns:
            pandas.DataFrame: A DataFrame with the value counts of payment plans
                              containing the word 'Zone'.

        Raises:
            ValueError: If the dataframe was not loaded successfully.
        """
        if self.df is None:
            raise ValueError("Dataframe not loaded. Cannot process data.")

        # 1. Filter for rows where 'Payment Plan Name' contains 'Zone'.
        #    `na=False` ensures that empty/NaN cells don't cause an error.
        zone_plans_df = self.df[
            self.df["Payment Plan Name"].str.contains("Zone", na=False)
        ]

        # 2. Get the value counts of the 'Payment Plan Name' column. This returns a pandas Series.
        series_result = zone_plans_df["Payment Plan Name"].value_counts()

        # 3. Convert the Series to a DataFrame, sort it, and give it clear column names.
        df_result = series_result.rename_axis("Payment Plan Name").reset_index(
            name="counts"
        )
        df_result.sort_values(by="Payment Plan Name", inplace=True)

        return df_result
