#Sequence:1
import pandas as pd

sis_online_input_file_for_Concatenate = 'SIS_Online_10_aug_2017.csv'
date = '10_aug_2017'

input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input" + "\\"  + sis_online_input_file_for_Concatenate
output_sis_online_for_concatenate_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_sis_online_for_concatenate_" + date + ".xlsx"
output_sis_online_for_concatenate_path_reuse = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input\selected_columns_sis_online_for_concatenate_" + date + ".xlsx"
output_sis_online_all_date_fields_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_sis_online_all_date_fields_" + date + ".xlsx"

sis_online = pd.read_csv(input_sis_online_path, parse_dates=True, low_memory=False)

delete_first_row = sis_online.ix[1:]
#print(delete_first_row)
#sis_online_selected_columns_for_concatenated = delete_first_row.iloc[:, (0, 2, 1, 9)] #same things does by the following line, but here with column index
sis_online_selected_columns_for_concatenated = delete_first_row[['formResultId', 'sis_cl_first_nm', 'sis_cl_last_nm', 'sis_track_num']]
sis_online_selected_columns_for_concatenated.rename(columns={'formResultId': 'SIS_ID', 'sis_cl_first_nm': 'First_Name', 'sis_cl_last_nm': 'Last_Name', 'sis_track_num': 'Tracking Number'}, inplace=True)
#print(sis_online_selected_columns_for_concatenated)

#sis_online_selected_columns_with_date_fields = delete_first_row.iloc[:, (0, 9, 12, 190, 13, 15, 14)]
sis_online_selected_columns_with_date_fields = delete_first_row[['formResultId', 'sis_track_num', 'statusText', 'locked', 'statusChangeDate', 'dateUpdated', 'sis_completed_dt']]
sis_online_selected_columns_with_date_fields.rename(columns={'formResultId': 'SIS_ID', 'sis_track_num': 'Tracking Number', 'dateUpdated': 'dateUpdated_SIS_Online', 'sis_completed_dt': 'Completed Date',}, inplace=True)
#sis_online_selected_columns_with_date_fields['locked'] = sis_online_selected_columns_with_date_fields['locked'].str.upper()
print(sis_online_selected_columns_with_date_fields[:5])

sis_online_selected_columns_for_concatenated.to_excel(output_sis_online_for_concatenate_path, index= False, sheet_name='Sheet1')
sis_online_selected_columns_for_concatenated.to_excel(output_sis_online_for_concatenate_path_reuse, index= False, sheet_name='Sheet1')
sis_online_selected_columns_with_date_fields.to_excel(output_sis_online_all_date_fields_path, index= False, sheet_name='Sheet1')

print("Column Count of SIS Online CSV: " + str(len(sis_online.columns)))
print('Successfull!')
