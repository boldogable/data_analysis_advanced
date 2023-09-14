import pandas as pd

class ExcelLoader:
    def load_data(self, file_paths):
        dataframes = []
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            dataframes.append(df)
        return dataframes
