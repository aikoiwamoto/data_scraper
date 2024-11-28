import pandas as pd

# Load your Excel dataset into a pandas DataFrame
df = pd.read_excel('INFS5135 - Master Data Sheet - Copy.xlsx')

# Group the data by 'Suburb' and 'Car Brand' and count the instances
result = df.groupby(['Suburb', 'Plate Type']).size().reset_index(name='Count')

# Save the result to a new Excel file
result.to_excel('plate_type_suburb_counts.xlsx', index=False)
