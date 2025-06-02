#Sequence:1
import pandas as pd

r_online_input_file_for_Concatenate = 'r_Online_10_aug_2017.csv'
date = '10_aug_2017'

input_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input" + "\\"  + r_online_input_file_for_Concatenate
output_r_online_for_concatenate_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_online_for_concatenate_" + date + ".xlsx"
output_r_online_for_concatenate_path_reuse = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input\selected_columns_r_online_for_concatenate_" + date + ".xlsx"
output_r_online_all_date_fields_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_online_all_date_fields_" + date + ".xlsx"

r_online = pd.read_csv(input_r_online_path, parse_dates=True, low_memory=False)

delete_first_row = r_online.ix[1:]
#print(delete_first_row)
#r_online_selected_columns_for_concatenated = delete_first_row.iloc[:, (0, 2, 1, 9)] #same things does by the following line, but here with column index
r_online_selected_columns_for_concatenated = delete_first_row[['formResultId', 'r_cl_first_nm', 'r_cl_last_nm', 'r_track_num']]
r_online_selected_columns_for_concatenated.rename(columns={'formResultId': 'r_ID', 'r_cl_first_nm': 'First_Name', 'r_cl_last_nm': 'Last_Name', 'r_track_num': 'Tracking Number'}, inplace=True)
#print(r_online_selected_columns_for_concatenated)

#r_online_selected_columns_with_date_fields = delete_first_row.iloc[:, (0, 9, 12, 190, 13, 15, 14)]
r_online_selected_columns_with_date_fields = delete_first_row[['formResultId', 'r_track_num', 'statusText', 'locked', 'statusChangeDate', 'dateUpdated', 'r_completed_dt']]
r_online_selected_columns_with_date_fields.rename(columns={'formResultId': 'r_ID', 'r_track_num': 'Tracking Number', 'dateUpdated': 'dateUpdated_r_Online', 'r_completed_dt': 'Completed Date',}, inplace=True)
#r_online_selected_columns_with_date_fields['locked'] = r_online_selected_columns_with_date_fields['locked'].str.upper()
print(r_online_selected_columns_with_date_fields[:5])

r_online_selected_columns_for_concatenated.to_excel(output_r_online_for_concatenate_path, index= False, sheet_name='Sheet1')
r_online_selected_columns_for_concatenated.to_excel(output_r_online_for_concatenate_path_reuse, index= False, sheet_name='Sheet1')
r_online_selected_columns_with_date_fields.to_excel(output_r_online_all_date_fields_path, index= False, sheet_name='Sheet1')

print("Column Count of r Online CSV: " + str(len(r_online.columns)))
print('Successfull!')
