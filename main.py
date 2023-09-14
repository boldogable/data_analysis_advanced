from excel_loader import ExcelLoader
from data_analyzer import DataAnalyzer

def main():
    # Initialize data loader
    excel_loader = ExcelLoader()

    # Load Excel files
    file_paths = ['file1.xlsx', 'file2.xlsx']  # Add your file paths
    dataframes = excel_loader.load_data(file_paths)

    # Initialize data analyzer
    data_analyzer = DataAnalyzer()

    # Analyze data and save results
    results = data_analyzer.analyze_data(dataframes)

    # Print results
    data_analyzer.print_results(results)

    # Save results
    data_analyzer.save_results(results, "calculated.xlsx")

    # Save highlighted results
    data_analyzer.save_highlighted_results(dataframes, "highlighted_results.xlsx")

if __name__ == "__main__":
    main()
