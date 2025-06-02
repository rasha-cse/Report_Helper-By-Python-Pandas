#Sequence:4 updated
import pandas as pd
import numpy as np

#r_input_file_with_date_fields = 'r_prod_with_all_date_fields_8_aug_2017_new(2).xlsx'
date = '8_sep_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_r_file_all_date_fields_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_sis_online_all_date_fields_" + date + ".xlsx"
output_exists_both_sides_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_both_with_date_fields_" + date + ".xlsx"


input_r_dataframe = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'])
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'])

exists_both_sides_dataframe = pd.merge(input_r_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
exists_both_sides_dataframe['Exist'] = np.where(exists_both_sides_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (exists_both_sides_dataframe[['SIS_ID', 'Tracking Number', 'Exist', 'dateUpdated_SIS_Online']])
#print(df['Exist'].unique())
print(exists_both_sides_dataframe['Exist'].value_counts())

######################################## SIS Online SIS_STATUS ###########################################
def sis_online_sis_status_formation(exists_both_sides_dataframe):
    if exists_both_sides_dataframe['locked'] == True:
        exists_both_sides_dataframe['SIS_STATUS'] = exists_both_sides_dataframe['statusText'] + '_LOCKED'
    else:
        exists_both_sides_dataframe['SIS_STATUS'] = exists_both_sides_dataframe['statusText']
    return exists_both_sides_dataframe['SIS_STATUS']

exists_both_sides_dataframe['SIS_STATUS_SIS_Online'] = exists_both_sides_dataframe.apply(sis_online_sis_status_formation, axis=1)

is_locked = exists_both_sides_dataframe['locked'] == True
not_locked = exists_both_sides_dataframe['locked'] == False
######################################## SIS Online SIS_STATUS ###########################################

exists_both = exists_both_sides_dataframe['Exist'] == True
#print(exists_both_sides_dataframe[exists_both][:5])
print(exists_both_sides_dataframe[exists_both]['Exist'].value_counts())
exists_both_sides = exists_both_sides_dataframe[exists_both][['SIS_ID', 'SIS_TRACKING_NUM', 'SIS_STATUS', 'SIS_STATUS_SIS_Online', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE', 'dateUpdated_SIS_Online']]

exists_both_sides["CREATED"] = exists_both_sides["CREATED"].map(lambda x: str(x)[:10])
exists_both_sides["UPDATED"] = exists_both_sides["UPDATED"].map(lambda x: str(x)[:10])
exists_both_sides["STATUS_CHANGE_DATE"] = exists_both_sides["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:10])
exists_both_sides["COMPLETED_DATE"] = exists_both_sides["COMPLETED_DATE"].map(lambda x: str(x)[:10])

exists_both_sides["dateUpdated_SIS_Online"] = pd.to_datetime(exists_both_sides["dateUpdated_SIS_Online"])
exists_both_sides["dateUpdated_SIS_Online"] = exists_both_sides["dateUpdated_SIS_Online"].dt.strftime("%m/%d/%Y")   #("%d-%b-%y")
#exists_both_sides["dateUpdated_SIS_Online"] = exists_both_sides["dateUpdated_SIS_Online"].map(lambda x: str(x)[:10])

exists_both_sides.rename(columns={'SIS_STATUS': 'SIS_STATUS_r', 'CREATED': 'CREATED_AT_r', 'UPDATED': 'UPDATED_AT_r', 'dateUpdated_SIS_Online': 'Date_Updated_SIS_Online'}, inplace=True)

exists_both_sides['Matched/Mismatched?'] = np.where(exists_both_sides['SIS_STATUS_r'] == exists_both_sides['SIS_STATUS_SIS_Online'], 'Matched', 'Mismatched')  #Matched/Mismatched?

exists_both_sides = exists_both_sides[['SIS_ID', 'SIS_TRACKING_NUM', 'SIS_STATUS_r', 'SIS_STATUS_SIS_Online', 'Matched/Mismatched?', 'CREATED_AT_r', 'UPDATED_AT_r', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE', 'Date_Updated_SIS_Online']]

writer_exists_both_sides = pd.ExcelWriter(output_exists_both_sides_path, engine='xlsxwriter')
exists_both_sides.to_excel(writer_exists_both_sides, index=False, sheet_name='Sheet1')
writer_exists_both_sides.save()
