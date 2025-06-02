#Sequence:3 updated
import pandas as pd
import numpy as np

date = '8_sep_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_r_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_sis_online_" + date + ".xlsx"
output_concatenated_comparison_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_comparison_" + date + ".xlsx"


input_r_dataframe = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })

#input_r_dataframe.rename(columns={'FIRST_NAME' : 'FIRST_NAME_r',  'LAST_NAME' : 'LAST_NAME_r', 'SIS_TRACKING_NUM': 'SIS_TRACKING_NUM_r'}, inplace=True)
input_r_dataframe = input_r_dataframe.add_suffix('_r')
#print(list(input_r_dataframe.columns.values))
input_r_dataframe.rename(columns={'SIS_ID_r' : 'SIS_ID', 'Concatenated_r_r': 'Concatenated_r'}, inplace=True)
#input_sis_online_dataframe.rename(columns={'First_Name' : 'First_Name_SIS_Online',  'Last_Name' : 'Last_Name_SIS_Online', 'Tracking Number': 'Tracking Number_SIS_Online'}, inplace=True)
input_sis_online_dataframe = input_sis_online_dataframe.add_suffix('_SIS_Online')
#print(list(input_sis_online_dataframe.columns.values))
input_sis_online_dataframe.rename(columns={'SIS_ID_SIS_Online' : 'SIS_ID', 'Concatenated_SIS_Online_SIS_Online': 'Concatenated_SIS_Online'}, inplace=True)

exists_both_sides_dataframe = pd.merge(input_r_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
exists_both_sides_dataframe['Exist'] = np.where(exists_both_sides_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (exists_both_sides_dataframe[['SIS_ID', 'Tracking Number', 'Exist', 'dateUpdated_SIS_Online']])
#print(df['Exist'].unique())
print(exists_both_sides_dataframe['Exist'].value_counts())

exists_both = exists_both_sides_dataframe['Exist'] == True
#print(exists_both_sides_dataframe[exists_both][:5])
print(exists_both_sides_dataframe[exists_both]['Exist'].value_counts())

exists_both_sides = exists_both_sides_dataframe[exists_both][['SIS_ID', 'FIRST_NAME_r', 'LAST_NAME_r', 'SIS_TRACKING_NUM_r', 'HOME_LIVING_RAW_r', 'HOME_LIVING_STANDARD_r', 'HOME_LIVING_PERCENTILE_r', 'COMMUNITY_LIVING_RAW_r', 'COMMUNITY_LIVING_STANDARD_r', 'COMMUNITY_LIVING_PERCENTILE_r', 'LIFELONG_LEARNING_RAW_r', 'LIFELONG_LEARNING_STANDARD_r', 'LIFELONG_LEARNING_PERCENTILE_r', 'EMPLOYMENT_RAW_r', 'EMPLOYMENT_STANDARD_r', 'EMPLOYMENT_PERCENTILE_r', 'HEALTH_SAFETY_RAW_r', 'HEALTH_SAFETY_STANDARD_r', 'HEALTH_SAFETY_PERCENTILE_r', 'SOCIAL_RAW_r', 'SOCIAL_STANDARD_r', 'SOCIAL_PERCENTILE_r', 'TOTAL_SCORE_NEEDED_INDEX_r', 'EXCEPTION_MEDICAL_TOTAL_r', 'EXCEPTION_BEHAVE_TOTAL_r', 'Concatenated_r', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'HOME_LIVING_RAW_SIS_Online', 'HOME_LIVING_STANDARD_SIS_Online', 'HOME_LIVING_PERCENTILE_SIS_Online', 'COMMUNITY_LIVING_RAW_SIS_Online', 'COMMUNITY_LIVING_STANDARD_SIS_Online', 'COMMUNITY_LIVING_PERCENTILE_SIS_Online', 'LIFELONG_LEARNING_RAW_SIS_Online', 'LIFELONG_LEARNING_STANDARD_SIS_Online', 'LIFELONG_LEARNING_PERCENTILE_SIS_Online', 'EMPLOYMENT_RAW_SIS_Online', 'EMPLOYMENT_STANDARD_SIS_Online', 'EMPLOYMENT_PERCENTILE_SIS_Online', 'HEALTH_SAFETY_RAW_SIS_Online', 'HEALTH_SAFETY_STANDARD_SIS_Online', 'HEALTH_SAFETY_PERCENTILE_SIS_Online', 'SOCIAL_RAW_SIS_Online', 'SOCIAL_STANDARD_SIS_Online', 'SOCIAL_PERCENTILE_SIS_Online', 'TOTAL_SCORE_NEEDED_INDEX_SIS_Online', 'EXCEPTION_MEDICAL_TOTAL_SIS_Online', 'EXCEPTION_BEHAVE_TOTAL_SIS_Online', 'Concatenated_SIS_Online']]

exists_both_sides['Concatenated_r'] = exists_both_sides['Concatenated_r'].str.upper()
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.upper()

exists_both_sides['FIRST_NAME_r'] = exists_both_sides['FIRST_NAME_r'].str.upper()
exists_both_sides['LAST_NAME_r'] = exists_both_sides['LAST_NAME_r'].str.upper()
exists_both_sides['FIRST_NAME_r'] = exists_both_sides['FIRST_NAME_r'].str.replace(' ', '')
exists_both_sides['LAST_NAME_r'] = exists_both_sides['LAST_NAME_r'].str.replace(' ', '')

exists_both_sides['First_Name_SIS_Online'] = exists_both_sides['First_Name_SIS_Online'].str.upper()
exists_both_sides['Last_Name_SIS_Online'] = exists_both_sides['Last_Name_SIS_Online'].str.upper()

exists_both_sides['EXCEPTION_MEDICAL_TOTAL_r'] = exists_both_sides['EXCEPTION_MEDICAL_TOTAL_r'].astype(str)
exists_both_sides['EXCEPTION_BEHAVE_TOTAL_r'] = exists_both_sides['EXCEPTION_BEHAVE_TOTAL_r'].astype(str)

exists_both_sides['Concatenated_r'] = exists_both_sides['Concatenated_r'].str.replace(' ', '')
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.replace(' ', '')


exists_both_sides['Matched/Mismatched?'] = np.where(exists_both_sides['Concatenated_r'] == exists_both_sides['Concatenated_SIS_Online'], 'Matched',
                                                    np.where(exists_both_sides['SIS_TRACKING_NUM_r'] != exists_both_sides['Tracking Number_SIS_Online'], 'Tracking # Mismatched',
                                                             np.where((exists_both_sides['FIRST_NAME_r'] + exists_both_sides['LAST_NAME_r']) != (exists_both_sides['First_Name_SIS_Online'] + exists_both_sides['Last_Name_SIS_Online']), 'Name Mismatched',
                                                                      np.where((exists_both_sides['EXCEPTION_MEDICAL_TOTAL_r'] + exists_both_sides['EXCEPTION_BEHAVE_TOTAL_r']) == '00', 'Exceptional & Behavioral Zeroes in r',
                                                                               np.where(exists_both_sides['EXCEPTION_MEDICAL_TOTAL_r'] == 'nan', 'NULL Scores in r', 'SIS Score Mismatch')))))
#exists_both_sides['Tracking # Matched/Mismatched?'] = np.where(exists_both_sides['SIS_TRACKING_NUM_r'] == exists_both_sides['Tracking Number_SIS_Online'], 'Matched', 'Mismatched')

exists_both_sides = exists_both_sides[['SIS_ID', 'FIRST_NAME_r', 'LAST_NAME_r', 'SIS_TRACKING_NUM_r', 'HOME_LIVING_RAW_r', 'HOME_LIVING_STANDARD_r', 'HOME_LIVING_PERCENTILE_r', 'COMMUNITY_LIVING_RAW_r', 'COMMUNITY_LIVING_STANDARD_r', 'COMMUNITY_LIVING_PERCENTILE_r', 'LIFELONG_LEARNING_RAW_r', 'LIFELONG_LEARNING_STANDARD_r', 'LIFELONG_LEARNING_PERCENTILE_r', 'EMPLOYMENT_RAW_r', 'EMPLOYMENT_STANDARD_r', 'EMPLOYMENT_PERCENTILE_r', 'HEALTH_SAFETY_RAW_r', 'HEALTH_SAFETY_STANDARD_r', 'HEALTH_SAFETY_PERCENTILE_r', 'SOCIAL_RAW_r', 'SOCIAL_STANDARD_r', 'SOCIAL_PERCENTILE_r', 'TOTAL_SCORE_NEEDED_INDEX_r', 'EXCEPTION_MEDICAL_TOTAL_r', 'EXCEPTION_BEHAVE_TOTAL_r', 'Concatenated_r', 'Matched/Mismatched?', 'Concatenated_SIS_Online', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'HOME_LIVING_RAW_SIS_Online', 'HOME_LIVING_STANDARD_SIS_Online', 'HOME_LIVING_PERCENTILE_SIS_Online', 'COMMUNITY_LIVING_RAW_SIS_Online', 'COMMUNITY_LIVING_STANDARD_SIS_Online', 'COMMUNITY_LIVING_PERCENTILE_SIS_Online', 'LIFELONG_LEARNING_RAW_SIS_Online', 'LIFELONG_LEARNING_STANDARD_SIS_Online', 'LIFELONG_LEARNING_PERCENTILE_SIS_Online', 'EMPLOYMENT_RAW_SIS_Online', 'EMPLOYMENT_STANDARD_SIS_Online', 'EMPLOYMENT_PERCENTILE_SIS_Online', 'HEALTH_SAFETY_RAW_SIS_Online', 'HEALTH_SAFETY_STANDARD_SIS_Online', 'HEALTH_SAFETY_PERCENTILE_SIS_Online', 'SOCIAL_RAW_SIS_Online', 'SOCIAL_STANDARD_SIS_Online', 'SOCIAL_PERCENTILE_SIS_Online', 'TOTAL_SCORE_NEEDED_INDEX_SIS_Online', 'EXCEPTION_MEDICAL_TOTAL_SIS_Online', 'EXCEPTION_BEHAVE_TOTAL_SIS_Online']]

writer_exists_both_sides = pd.ExcelWriter(output_concatenated_comparison_path, engine='xlsxwriter')
exists_both_sides.to_excel(writer_exists_both_sides, index=False, sheet_name='Sheet1')
writer_exists_both_sides.save()
