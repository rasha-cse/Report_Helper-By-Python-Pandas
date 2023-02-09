#Sequence:3
import pandas as pd
import numpy as np

date = '8_aug_2017'

input_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_therap_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_sis_online_" + date + ".xlsx"
output_concatenated_comparison_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_comparison_" + date + ".xlsx"


input_therap_dataframe = pd.read_excel(input_therap_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x)})
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'], converters={'Tracking Number': lambda x: str(x)})

input_therap_dataframe.rename(columns={'FIRST_NAME' : 'FIRST_NAME_THERAP',  'LAST_NAME' : 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM': 'SIS_TRACKING_NUM_THERAP'}, inplace=True)
input_sis_online_dataframe.rename(columns={'First_Name' : 'First_Name_SIS_Online',  'Last_Name' : 'Last_Name_SIS_Online', 'Tracking Number': 'Tracking Number_SIS_Online'}, inplace=True)

exists_both_sides_dataframe = pd.merge(input_therap_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
exists_both_sides_dataframe['Exist'] = np.where(exists_both_sides_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (exists_both_sides_dataframe[['SIS_ID', 'Tracking Number', 'Exist', 'dateUpdated_SIS_Online']])
#print(df['Exist'].unique())
print(exists_both_sides_dataframe['Exist'].value_counts())

exists_both = exists_both_sides_dataframe['Exist'] == True
#print(exists_both_sides_dataframe[exists_both][:5])
print(exists_both_sides_dataframe[exists_both]['Exist'].value_counts())

exists_both_sides = exists_both_sides_dataframe[exists_both][['SIS_ID', 'FIRST_NAME_THERAP', 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM_THERAP', 'Concatenated_Therap', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'Concatenated_SIS_Online']]

exists_both_sides['Concatenated_Therap'] = exists_both_sides['Concatenated_Therap'].str.upper()
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.upper()

exists_both_sides['Concatenated_Therap'] = exists_both_sides['Concatenated_Therap'].str.replace(' ', '')
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.replace(' ', '')


exists_both_sides['Matched/Mismatched?'] = np.where(exists_both_sides['Concatenated_Therap'] == exists_both_sides['Concatenated_SIS_Online'], 'Matched',
                                                    np.where(exists_both_sides['SIS_TRACKING_NUM_THERAP'] == exists_both_sides['Tracking Number_SIS_Online'], 'Name Mismatched', 'Tracking # Mismatched'))
#exists_both_sides['Tracking # Matched/Mismatched?'] = np.where(exists_both_sides['SIS_TRACKING_NUM_THERAP'] == exists_both_sides['Tracking Number_SIS_Online'], 'Matched', 'Mismatched')

exists_both_sides = exists_both_sides[['SIS_ID', 'FIRST_NAME_THERAP', 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM_THERAP', 'Concatenated_Therap', 'Matched/Mismatched?', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'Concatenated_SIS_Online']]

writer_exists_both_sides = pd.ExcelWriter(output_concatenated_comparison_path, engine='xlsxwriter')
exists_both_sides.to_excel(writer_exists_both_sides, index=False, sheet_name='Sheet1')
writer_exists_both_sides.save()