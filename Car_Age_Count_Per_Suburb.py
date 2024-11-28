import pandas as pd

# Load your Excel dataset into a pandas DataFrame
df = pd.read_excel('INFS5135 - Master Data Sheet - Copy.xlsx')

# Group the data by 'Suburb' and 'Car Brand' and count the instances
result = df.groupby(['Suburb', 'Car Age (Current Year +1 - Model Year)']).size().reset_index(name='Count')

# Save the result to a new Excel file
result.to_excel('car_age_suburb_counts.xlsx', index=False)