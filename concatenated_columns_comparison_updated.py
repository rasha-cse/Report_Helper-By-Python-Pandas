#Sequence:3 updated
import pandas as pd
import numpy as np

date = '8_sep_2017'

input_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_therap_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_sis_online_" + date + ".xlsx"
output_concatenated_comparison_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_comparison_" + date + ".xlsx"


input_therap_dataframe = pd.read_excel(input_therap_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })

#input_therap_dataframe.rename(columns={'FIRST_NAME' : 'FIRST_NAME_THERAP',  'LAST_NAME' : 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM': 'SIS_TRACKING_NUM_THERAP'}, inplace=True)
input_therap_dataframe = input_therap_dataframe.add_suffix('_THERAP')
#print(list(input_therap_dataframe.columns.values))
input_therap_dataframe.rename(columns={'SIS_ID_THERAP' : 'SIS_ID', 'Concatenated_Therap_THERAP': 'Concatenated_Therap'}, inplace=True)
#input_sis_online_dataframe.rename(columns={'First_Name' : 'First_Name_SIS_Online',  'Last_Name' : 'Last_Name_SIS_Online', 'Tracking Number': 'Tracking Number_SIS_Online'}, inplace=True)
input_sis_online_dataframe = input_sis_online_dataframe.add_suffix('_SIS_Online')
#print(list(input_sis_online_dataframe.columns.values))
input_sis_online_dataframe.rename(columns={'SIS_ID_SIS_Online' : 'SIS_ID', 'Concatenated_SIS_Online_SIS_Online': 'Concatenated_SIS_Online'}, inplace=True)

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

exists_both_sides = exists_both_sides_dataframe[exists_both][['SIS_ID', 'FIRST_NAME_THERAP', 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM_THERAP', 'HOME_LIVING_RAW_THERAP', 'HOME_LIVING_STANDARD_THERAP', 'HOME_LIVING_PERCENTILE_THERAP', 'COMMUNITY_LIVING_RAW_THERAP', 'COMMUNITY_LIVING_STANDARD_THERAP', 'COMMUNITY_LIVING_PERCENTILE_THERAP', 'LIFELONG_LEARNING_RAW_THERAP', 'LIFELONG_LEARNING_STANDARD_THERAP', 'LIFELONG_LEARNING_PERCENTILE_THERAP', 'EMPLOYMENT_RAW_THERAP', 'EMPLOYMENT_STANDARD_THERAP', 'EMPLOYMENT_PERCENTILE_THERAP', 'HEALTH_SAFETY_RAW_THERAP', 'HEALTH_SAFETY_STANDARD_THERAP', 'HEALTH_SAFETY_PERCENTILE_THERAP', 'SOCIAL_RAW_THERAP', 'SOCIAL_STANDARD_THERAP', 'SOCIAL_PERCENTILE_THERAP', 'TOTAL_SCORE_NEEDED_INDEX_THERAP', 'EXCEPTION_MEDICAL_TOTAL_THERAP', 'EXCEPTION_BEHAVE_TOTAL_THERAP', 'Concatenated_Therap', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'HOME_LIVING_RAW_SIS_Online', 'HOME_LIVING_STANDARD_SIS_Online', 'HOME_LIVING_PERCENTILE_SIS_Online', 'COMMUNITY_LIVING_RAW_SIS_Online', 'COMMUNITY_LIVING_STANDARD_SIS_Online', 'COMMUNITY_LIVING_PERCENTILE_SIS_Online', 'LIFELONG_LEARNING_RAW_SIS_Online', 'LIFELONG_LEARNING_STANDARD_SIS_Online', 'LIFELONG_LEARNING_PERCENTILE_SIS_Online', 'EMPLOYMENT_RAW_SIS_Online', 'EMPLOYMENT_STANDARD_SIS_Online', 'EMPLOYMENT_PERCENTILE_SIS_Online', 'HEALTH_SAFETY_RAW_SIS_Online', 'HEALTH_SAFETY_STANDARD_SIS_Online', 'HEALTH_SAFETY_PERCENTILE_SIS_Online', 'SOCIAL_RAW_SIS_Online', 'SOCIAL_STANDARD_SIS_Online', 'SOCIAL_PERCENTILE_SIS_Online', 'TOTAL_SCORE_NEEDED_INDEX_SIS_Online', 'EXCEPTION_MEDICAL_TOTAL_SIS_Online', 'EXCEPTION_BEHAVE_TOTAL_SIS_Online', 'Concatenated_SIS_Online']]

exists_both_sides['Concatenated_Therap'] = exists_both_sides['Concatenated_Therap'].str.upper()
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.upper()

exists_both_sides['FIRST_NAME_THERAP'] = exists_both_sides['FIRST_NAME_THERAP'].str.upper()
exists_both_sides['LAST_NAME_THERAP'] = exists_both_sides['LAST_NAME_THERAP'].str.upper()
exists_both_sides['FIRST_NAME_THERAP'] = exists_both_sides['FIRST_NAME_THERAP'].str.replace(' ', '')
exists_both_sides['LAST_NAME_THERAP'] = exists_both_sides['LAST_NAME_THERAP'].str.replace(' ', '')

exists_both_sides['First_Name_SIS_Online'] = exists_both_sides['First_Name_SIS_Online'].str.upper()
exists_both_sides['Last_Name_SIS_Online'] = exists_both_sides['Last_Name_SIS_Online'].str.upper()

exists_both_sides['EXCEPTION_MEDICAL_TOTAL_THERAP'] = exists_both_sides['EXCEPTION_MEDICAL_TOTAL_THERAP'].astype(str)
exists_both_sides['EXCEPTION_BEHAVE_TOTAL_THERAP'] = exists_both_sides['EXCEPTION_BEHAVE_TOTAL_THERAP'].astype(str)

exists_both_sides['Concatenated_Therap'] = exists_both_sides['Concatenated_Therap'].str.replace(' ', '')
exists_both_sides['Concatenated_SIS_Online'] = exists_both_sides['Concatenated_SIS_Online'].str.replace(' ', '')


exists_both_sides['Matched/Mismatched?'] = np.where(exists_both_sides['Concatenated_Therap'] == exists_both_sides['Concatenated_SIS_Online'], 'Matched',
                                                    np.where(exists_both_sides['SIS_TRACKING_NUM_THERAP'] != exists_both_sides['Tracking Number_SIS_Online'], 'Tracking # Mismatched',
                                                             np.where((exists_both_sides['FIRST_NAME_THERAP'] + exists_both_sides['LAST_NAME_THERAP']) != (exists_both_sides['First_Name_SIS_Online'] + exists_both_sides['Last_Name_SIS_Online']), 'Name Mismatched',
                                                                      np.where((exists_both_sides['EXCEPTION_MEDICAL_TOTAL_THERAP'] + exists_both_sides['EXCEPTION_BEHAVE_TOTAL_THERAP']) == '00', 'Exceptional & Behavioral Zeroes in Therap',
                                                                               np.where(exists_both_sides['EXCEPTION_MEDICAL_TOTAL_THERAP'] == 'nan', 'NULL Scores in Therap', 'SIS Score Mismatch')))))
#exists_both_sides['Tracking # Matched/Mismatched?'] = np.where(exists_both_sides['SIS_TRACKING_NUM_THERAP'] == exists_both_sides['Tracking Number_SIS_Online'], 'Matched', 'Mismatched')

exists_both_sides = exists_both_sides[['SIS_ID', 'FIRST_NAME_THERAP', 'LAST_NAME_THERAP', 'SIS_TRACKING_NUM_THERAP', 'HOME_LIVING_RAW_THERAP', 'HOME_LIVING_STANDARD_THERAP', 'HOME_LIVING_PERCENTILE_THERAP', 'COMMUNITY_LIVING_RAW_THERAP', 'COMMUNITY_LIVING_STANDARD_THERAP', 'COMMUNITY_LIVING_PERCENTILE_THERAP', 'LIFELONG_LEARNING_RAW_THERAP', 'LIFELONG_LEARNING_STANDARD_THERAP', 'LIFELONG_LEARNING_PERCENTILE_THERAP', 'EMPLOYMENT_RAW_THERAP', 'EMPLOYMENT_STANDARD_THERAP', 'EMPLOYMENT_PERCENTILE_THERAP', 'HEALTH_SAFETY_RAW_THERAP', 'HEALTH_SAFETY_STANDARD_THERAP', 'HEALTH_SAFETY_PERCENTILE_THERAP', 'SOCIAL_RAW_THERAP', 'SOCIAL_STANDARD_THERAP', 'SOCIAL_PERCENTILE_THERAP', 'TOTAL_SCORE_NEEDED_INDEX_THERAP', 'EXCEPTION_MEDICAL_TOTAL_THERAP', 'EXCEPTION_BEHAVE_TOTAL_THERAP', 'Concatenated_Therap', 'Matched/Mismatched?', 'Concatenated_SIS_Online', 'First_Name_SIS_Online', 'Last_Name_SIS_Online', 'Tracking Number_SIS_Online', 'HOME_LIVING_RAW_SIS_Online', 'HOME_LIVING_STANDARD_SIS_Online', 'HOME_LIVING_PERCENTILE_SIS_Online', 'COMMUNITY_LIVING_RAW_SIS_Online', 'COMMUNITY_LIVING_STANDARD_SIS_Online', 'COMMUNITY_LIVING_PERCENTILE_SIS_Online', 'LIFELONG_LEARNING_RAW_SIS_Online', 'LIFELONG_LEARNING_STANDARD_SIS_Online', 'LIFELONG_LEARNING_PERCENTILE_SIS_Online', 'EMPLOYMENT_RAW_SIS_Online', 'EMPLOYMENT_STANDARD_SIS_Online', 'EMPLOYMENT_PERCENTILE_SIS_Online', 'HEALTH_SAFETY_RAW_SIS_Online', 'HEALTH_SAFETY_STANDARD_SIS_Online', 'HEALTH_SAFETY_PERCENTILE_SIS_Online', 'SOCIAL_RAW_SIS_Online', 'SOCIAL_STANDARD_SIS_Online', 'SOCIAL_PERCENTILE_SIS_Online', 'TOTAL_SCORE_NEEDED_INDEX_SIS_Online', 'EXCEPTION_MEDICAL_TOTAL_SIS_Online', 'EXCEPTION_BEHAVE_TOTAL_SIS_Online']]

writer_exists_both_sides = pd.ExcelWriter(output_concatenated_comparison_path, engine='xlsxwriter')
exists_both_sides.to_excel(writer_exists_both_sides, index=False, sheet_name='Sheet1')
writer_exists_both_sides.save()