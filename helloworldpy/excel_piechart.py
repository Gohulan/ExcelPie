import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook

# Load the Excel file
filename = r'C:\test\Sales.xlsx'
wb = load_workbook(filename)

# Select the worksheet
ws = wb['Sales_Data']

# Read the data into a Pandas dataframe
df = pd.DataFrame(ws.values)

# Get the column names from the first row
df.columns = df.iloc[0]

# Drop the first row (column names) and any rows with NaN values
df = df.iloc[1:].dropna()

# Calculate the total amount
df['Total Amount'] = df['Quantity'] * df['Amount']

fig = plt.figure()
fig.suptitle('Location A')

# Create a pie chart
plt.pie(df['Total Amount'], labels=df['Sale Item'], autopct='%1.1f%%')
plt.title('Sales Summary')

# Show the chart
plt.show()
