#Sequence:3
import pandas as pd
import numpy as np

date = '8_aug_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\concatenated_r_" + date + ".xlsx"
input_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\concatenated_r_online_" + date + ".xlsx"
output_concatenated_comparison_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\concatenated_comparison_" + date + ".xlsx"


input_r_dataframe = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'], converters={'r_TRACKING_NUM': lambda x: str(x)})
input_r_online_dataframe = pd.read_excel(input_r_online_path, sheetname='Sheet1', na_values=['NA'], converters={'Tracking Number': lambda x: str(x)})

input_r_dataframe.rename(columns={'FIRST_NAME' : 'FIRST_NAME_r',  'LAST_NAME' : 'LAST_NAME_r', 'r_TRACKING_NUM': 'r_TRACKING_NUM_r'}, inplace=True)
input_r_online_dataframe.rename(columns={'First_Name' : 'First_Name_r_Online',  'Last_Name' : 'Last_Name_r_Online', 'Tracking Number': 'Tracking Number_r_Online'}, inplace=True)

exists_both_sides_dataframe = pd.merge(input_r_dataframe, input_r_online_dataframe, on=['r_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
exists_both_sides_dataframe['Exist'] = np.where(exists_both_sides_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (exists_both_sides_dataframe[['r_ID', 'Tracking Number', 'Exist', 'dateUpdated_r_Online']])
#print(df['Exist'].unique())
print(exists_both_sides_dataframe['Exist'].value_counts())

exists_both = exists_both_sides_dataframe['Exist'] == True
#print(exists_both_sides_dataframe[exists_both][:5])
print(exists_both_sides_dataframe[exists_both]['Exist'].value_counts())

exists_both_sides = exists_both_sides_dataframe[exists_both][['r_ID', 'FIRST_NAME_r', 'LAST_NAME_r', 'r_TRACKING_NUM_r', 'Concatenated_r', 'First_Name_r_Online', 'Last_Name_r_Online', 'Tracking Number_r_Online', 'Concatenated_r_Online']]

exists_both_sides['Concatenated_r'] = exists_both_sides['Concatenated_r'].str.upper()
exists_both_sides['Concatenated_r_Online'] = exists_both_sides['Concatenated_r_Online'].str.upper()

exists_both_sides['Concatenated_r'] = exists_both_sides['Concatenated_r'].str.replace(' ', '')
exists_both_sides['Concatenated_r_Online'] = exists_both_sides['Concatenated_r_Online'].str.replace(' ', '')


exists_both_sides['Matched/Mismatched?'] = np.where(exists_both_sides['Concatenated_r'] == exists_both_sides['Concatenated_r_Online'], 'Matched',
                                                    np.where(exists_both_sides['r_TRACKING_NUM_r'] == exists_both_sides['Tracking Number_r_Online'], 'Name Mismatched', 'Tracking # Mismatched'))
#exists_both_sides['Tracking # Matched/Mismatched?'] = np.where(exists_both_sides['r_TRACKING_NUM_r'] == exists_both_sides['Tracking Number_r_Online'], 'Matched', 'Mismatched')

exists_both_sides = exists_both_sides[['r_ID', 'FIRST_NAME_r', 'LAST_NAME_r', 'r_TRACKING_NUM_r', 'Concatenated_r', 'Matched/Mismatched?', 'First_Name_r_Online', 'Last_Name_r_Online', 'Tracking Number_r_Online', 'Concatenated_r_Online']]

writer_exists_both_sides = pd.ExcelWriter(output_concatenated_comparison_path, engine='xlsxwriter')
exists_both_sides.to_excel(writer_exists_both_sides, index=False, sheet_name='Sheet1')
writer_exists_both_sides.save()
