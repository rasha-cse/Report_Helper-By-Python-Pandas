import pandas as pd

input_unsorted_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input\Both_r_Online.xlsx"
output_sorted_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\sorted_Both_r_Online_20_apr_2017.xlsx"

input_dataframe = pd.read_excel(input_unsorted_path, sheetname='Sheet1', na_values=['NA'])

sorted_dataframe = input_dataframe.sort_values(['r_ID'], ascending=True)

writer_r = pd.ExcelWriter(output_sorted_path, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
sorted_dataframe.to_excel(writer_r, index=False, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer_r.save()

print('Sort Successfull!')
