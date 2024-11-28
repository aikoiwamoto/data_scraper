# from openpyxl import load_workbook

# excel_file = 'INFS5135_-_VIN_Data_Sheet.xlsx'
# sheet_name = 'Data Sheet'
# wb = load_workbook(excel_file)
# ws = wb[sheet_name]

# year_from_VIN = {
#     'A': [1980, 2010], 
#     'B': [1981, 2011],
#     'C': [1982, 2012],
#     'D': [1983, 2013],
#     'E': [1984, 2014],
#     'F': [1985, 2015],
#     'G': [1986, 2016],
#     'H': [1987, 2017],
#     'J': [1988, 2018],
#     'K': [1989, 2019],
#     'L': [1990, 2020],
#     'M': [1991, 2021],
#     'N': [1992, 2022],
#     'P': [1993, 2023],
#     'R': [1994, 2024],
#     'S': [1995, 2025],
#     'T': [1996, 2026],
#     'V': [1997, 2027],
#     'W': [1998, 2028],
#     'X': [1999, 2029],
#     'Y': 2000,
#     1: 2001, 
#     2: 2002, 
#     3: 2003, 
#     4: 2004, 
#     5: 2005, 
#     6: 2006,
#     7: 2007,
#     8: 2008,
#     9: 2009,
#     0: "The year for this car cannot be found."
# }


import pandas as pd

year_from_VIN = {
    'A': [1980, 2010], 'B': [1981, 2011], 'C': [1982, 2012], 'D': [1983, 2013], 'E': [1984, 2014],
    'F': [1985, 2015], 'G': [1986, 2016], 'H': [1987, 2017], 'J': [1988, 2018], 'K': [1989, 2019],
    'L': [1990, 2020], 'M': [1991, 2021], 'N': [1992, 2022], 'P': [1993, 2023], 'R': [1994, 2024],
    'S': [1995, 2025], 'T': [1996, 2026], 'V': [1997, 2027], 'W': [1998, 2028], 'X': [1999, 2029],
    'Y': 2000, 1: 2001, 2: 2002, 3: 2003, 4: 2004, 5: 2005, 6: 2006, 7: 2007, 8: 2008, 9: 2009, 0: "The year for this car cannot be found."
}


def determine_year_from_excel(file_path, sheet_name, key_column, start_year_column, end_year_column, model_year_column):
    # Read data from the specified sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Determine the keys from the key column
    keys = df[key_column]
    
    # Determine the model year based on the dictionary and the keys
    determined_years = []
    for key in keys:
        determined_year = "Year not found."
        if key in year_from_VIN:
            year_range = year_from_VIN[key]
            if isinstance(year_range, list):
                start, end = year_range
                determined_year = end if end <= 2024 else start
            elif isinstance(year_range, int):
                determined_year = year_range
        determined_years.append(determined_year)
    
    # Update the DataFrame in the specified model year column (column Z)
    df[model_year_column] = determined_years
    
    return df


# Example usage
file_path = "INFS5135_-_VIN_Data_Sheet.xlsx"
sheet_name = 'Scraped Data'  # Replace 'Sheet1' with the actual sheet name in your Excel file
key_column = "Model Year Code"  # Replace with your actual column name for dictionary keys
start_year_column = "Full Start Year"  # Replace with your actual column name for start year
end_year_column = "Full End Year"  # Replace with your actual column name for end year
model_year_column = "Model Year"  # Replace with your actual column name for input values

result_df = determine_year_from_excel(file_path, sheet_name, key_column, start_year_column, end_year_column, model_year_column)
print(result_df)
result_df = determine_year_from_excel(file_path, sheet_name, key_column, start_year_column, end_year_column, model_year_column)
result_df.to_excel("output_path.xlsx", index=False)  # Save the updated DataFrame to a new Excel file