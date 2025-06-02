#Sequence:1 updated
import pandas as pd

r_online_input_file_for_Concatenate = 'r_Online_8_sep_2017_prod.csv'
r_file_for_concatenation = 'r_all_scrores_dates_8_sep_2017_production.xlsx'
date = '8_sep_2017'

input_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input" + "\\"  + r_online_input_file_for_Concatenate
input_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input" + "\\" + r_file_for_concatenation

output_r_online_for_concatenate_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_online_for_concatenate_" + date + ".xlsx"
output_r_file_for_concatenate_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_file_for_concatenate_" + date + ".xlsx"
output_r_online_for_concatenate_path_reuse = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input\selected_columns_r_online_for_concatenate_" + date + ".xlsx"
output_r_online_all_date_fields_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_online_all_date_fields_" + date + ".xlsx"
output_r_with_all_date_fields_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_file_all_date_fields_" + date + ".xlsx"

r_online = pd.read_csv(input_r_online_path, parse_dates=True, low_memory=False)
r_file = pd.read_excel(input_r_path)

delete_first_row = r_online.ix[1:]
#print(delete_first_row)
#r_online_selected_columns_for_concatenated = delete_first_row.iloc[:, (0, 2, 1, 9)] #same things does by the following line, but here with column index
r_online_selected_columns_for_concatenated = delete_first_row[['formResultId', 'r_track_num', 'r_cl_first_nm', 'r_cl_last_nm', 'scr_2A_raw', 'scr_2A_std', 'scr_2A_pct', 'scr_2B_raw', 'scr_2B_std', 'scr_2B_pct', 'scr_2C_raw', 'scr_2C_std', 'scr_2C_pct', 'scr_2D_raw', 'scr_2D_std', 'scr_2D_pct', 'scr_2E_raw', 'scr_2E_std', 'scr_2E_pct', 'scr_2F_raw', 'scr_2F_std', 'scr_2F_pct', 'scr_support_needs_index', 'scr_1A_raw_total', 'scr_1B_raw_total']]
r_online_selected_columns_for_concatenated.rename(columns={'formResultId': 'r_ID', 'r_track_num': 'Tracking Number', 'r_cl_first_nm': 'First_Name', 'r_cl_last_nm': 'Last_Name', 'scr_2A_raw': 'HOME_LIVING_RAW', 'scr_2A_std': 'HOME_LIVING_STANDARD', 'scr_2A_pct': 'HOME_LIVING_PERCENTILE', 'scr_2B_raw': 'COMMUNITY_LIVING_RAW', 'scr_2B_std': 'COMMUNITY_LIVING_STANDARD', 'scr_2B_pct': 'COMMUNITY_LIVING_PERCENTILE', 'scr_2C_raw': 'LIFELONG_LEARNING_RAW', 'scr_2C_std': 'LIFELONG_LEARNING_STANDARD', 'scr_2C_pct': 'LIFELONG_LEARNING_PERCENTILE', 'scr_2D_raw': 'EMPLOYMENT_RAW', 'scr_2D_std': 'EMPLOYMENT_STANDARD', 'scr_2D_pct': 'EMPLOYMENT_PERCENTILE', 'scr_2E_raw': 'HEALTH_SAFETY_RAW', 'scr_2E_std': 'HEALTH_SAFETY_STANDARD', 'scr_2E_pct': 'HEALTH_SAFETY_PERCENTILE', 'scr_2F_raw': 'SOCIAL_RAW', 'scr_2F_std': 'SOCIAL_STANDARD', 'scr_2F_pct': 'SOCIAL_PERCENTILE', 'scr_support_needs_index': 'TOTAL_SCORE_NEEDED_INDEX', 'scr_1A_raw_total': 'EXCEPTION_MEDICAL_TOTAL', 'scr_1B_raw_total': 'EXCEPTION_BEHAVE_TOTAL'}, inplace=True)
print(r_online_selected_columns_for_concatenated[:5])

r_file_selected_columns_for_concatenated = r_file.drop(['r_STATUS', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE'], axis=1)
#r_file_selected_columns_for_concatenated = r_file_selected_columns_for_concatenated[['r_ID', 'FIRST_NAME',  'LAST_NAME', 'r_TRACKING_NUM' ]]
print(r_file_selected_columns_for_concatenated[:5])


#r_online_selected_columns_with_date_fields = delete_first_row.iloc[:, (0, 9, 12, 190, 13, 15, 14)]
r_online_selected_columns_with_date_fields = delete_first_row[['formResultId', 'r_track_num', 'statusText', 'locked', 'statusChangeDate', 'dateUpdated', 'r_completed_dt']]
r_online_selected_columns_with_date_fields.rename(columns={'formResultId': 'r_ID', 'r_track_num': 'Tracking Number', 'dateUpdated': 'dateUpdated_r_Online', 'r_completed_dt': 'Completed Date',}, inplace=True)
#r_online_selected_columns_with_date_fields['locked'] = r_online_selected_columns_with_date_fields['locked'].str.upper()
#print(r_online_selected_columns_with_date_fields[:5])


r_file_selected_columns_with_date_fields = r_file[['r_ID', 'r_TRACKING_NUM', 'r_STATUS', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE']]
#print(r_file_selected_columns_with_date_fields[:5])

r_online_selected_columns_for_concatenated.to_excel(output_r_online_for_concatenate_path, index= False, sheet_name='Sheet1')
r_file_selected_columns_for_concatenated.to_excel(output_r_file_for_concatenate_path, index= False, sheet_name='Sheet1')
r_online_selected_columns_for_concatenated.to_excel(output_r_online_for_concatenate_path_reuse, index= False, sheet_name='Sheet1')
r_online_selected_columns_with_date_fields.to_excel(output_r_online_all_date_fields_path, index= False, sheet_name='Sheet1')
r_file_selected_columns_with_date_fields.to_excel(output_r_with_all_date_fields_path, index= False, sheet_name='Sheet1')

print("Column Count of r Online CSV: " + str(len(r_online.columns)))
print('Successfull!')
