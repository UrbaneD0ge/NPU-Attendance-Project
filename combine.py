import pandas as pd
import glob

attFolder = r'F:\\OneDrive\\Documents\\CoA NPU Coordinator\\City Of Atlanta\\NPU - Documents\\2024\\NPU Attendance and Voting Reports\\TestData\\*.xlsx'
excel_files = glob.glob(attFolder)

df1 = pd.DataFrame()

for excel_file in excel_files:
  # drop the first 2 rows from each file and drop any columns beyond the first two
  df2 = pd.read_excel(excel_file).iloc[3:, :2]
  # remove any row that contains '@atlantaga.gov'
  # df2['Topic'] = df2['Topic'].astype(str)
  # cast all values to string
  df2 = df2.astype(str)
  df2 = df2[~df2['Topic'].str.contains('@atlantaga.gov')]
  # replace any NaN values with the value to its left
  # **not working as expected**
  df2 = df2.ffill(axis="columns")
  df1 = pd.concat([df1,df2], ignore_index=True)

df1.to_excel(r'F:\\OneDrive\\Documents\\CoA NPU Coordinator\\City Of Atlanta\NPU - Documents\\2024\\NPU Attendance and Voting Reports\\TestData\\Combined.xlsx', index=False)