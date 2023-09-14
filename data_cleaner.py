import pandas as pd

class DataCleaner:
    def clean_data(self, dataframes):
        cleaned_dataframes = []

        for df in dataframes:
            # Clean column names by removing leading/trailing spaces and special characters
            df.columns = df.columns.str.strip().str.replace(r'[^a-zA-Z0-9_]', '')

            # Clean data in each column
            for column in df.columns:
                if df[column].dtype == 'object':  # Only clean string columns
                    df[column] = df[column].str.strip()  # Remove leading/trailing spaces

            # Ensure all dataframes have the same columns (you can customize this logic)
            expected_columns = ['SupplierName', 'Inventory', 'Price', 'ProductNumber']
            missing_columns = set(expected_columns) - set(df.columns)
            for column in missing_columns:
                df[column] = None  # Add missing columns with None values

            cleaned_dataframes.append(df)

        return cleaned_dataframes
