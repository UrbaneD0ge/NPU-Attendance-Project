import pandas as pd
import glob, os

attFolder = r'.\\Attendance Data\\'
excel_files = glob.glob(attFolder)

df1 = pd.DataFrame()

# walk through the attFolder and combine all the files into one
for root, dirs, files in os.walk(attFolder):
    for file in files:
        if file.endswith('.xlsx'):
            # drop the first 2 rows from each file and drop any columns beyond the first two
            df2 = pd.read_excel(os.path.join(root, file)).iloc[3:, :2]
            df2 = df2.astype(str)
            # remove any row that contains '@atlantaga.gov'
            df2 = df2[~df2['Topic'].str.contains('@atlantaga.gov')]
            df1 = pd.concat([df1, df2], ignore_index=True)


# rename the second column to 'Email'
df1.rename(columns={df1.columns[0]: 'Name'}, inplace=True)
df1.rename(columns={df1.columns[1]: 'Email'}, inplace=True)

# TODO: NOT WORKING
# # fill forward any nan values in the 'Email' column using the value in the 'Name' column
# df1['Email'] = df1['Email'].fillna(method='ffill')

# count the number of times each email appears in the dataframe and add it to a new column
df1['Count'] = df1.groupby('Email')['Email'].transform('count')

# drop any duplicate rows
df1 = df1.drop_duplicates()

# sort the dataframe by the 'Count' column in descending order
df1 = df1.sort_values(by='Count', ascending=False)

# highlight any rows where the 'Count' column is greater than 3 in light green
def highlight_greater_than_3(s):
    return ['background-color: lightgreen' if v > 3 else '' for v in s]

df1.style.apply(highlight_greater_than_3, subset=['Count'])

df1.to_excel(r'Combined.xlsx', index=False)