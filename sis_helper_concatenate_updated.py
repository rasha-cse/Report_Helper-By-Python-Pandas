#Sequence:2 updated
import pandas as pd

#therap_input_file_for_Concatenate = 'Therap_for_Concatenate_8_aug_2017.xlsx'
date = '8_sep_2017'

input_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_therap_file_for_concatenate_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input\selected_columns_sis_online_for_concatenate_" + date + ".xlsx"
output_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_therap_" + date + ".xlsx"
output_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_sis_online_" + date + ".xlsx"

therap_prod = pd.read_excel(input_therap_path, sheetname='Sheet1', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })
sis_online = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'], converters={'Tracking Number': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })

#therap_prod['Concatenated_Therap'] = therap_prod.apply(lambda x:'%s%s%s%s' % (x['SIS_ID'],x['FIRST_NAME'],x['LAST_NAME'],x['SIS_TRACKING_NUM']),axis=1)
therap_prod['Concatenated_Therap'] = therap_prod['SIS_ID'].astype(str) + therap_prod['FIRST_NAME'] + therap_prod['LAST_NAME'] + therap_prod['SIS_TRACKING_NUM'].astype(str) + therap_prod['HOME_LIVING_RAW'].astype(str) + therap_prod['HOME_LIVING_STANDARD'].astype(str) + therap_prod['HOME_LIVING_PERCENTILE'].astype(str) + therap_prod['COMMUNITY_LIVING_RAW'].astype(str) + therap_prod['COMMUNITY_LIVING_STANDARD'].astype(str) + therap_prod['COMMUNITY_LIVING_PERCENTILE'].astype(str) + therap_prod['LIFELONG_LEARNING_RAW'].astype(str) + therap_prod['LIFELONG_LEARNING_STANDARD'].astype(str) + therap_prod['LIFELONG_LEARNING_PERCENTILE'].astype(str) + therap_prod['EMPLOYMENT_RAW'].astype(str) + therap_prod['EMPLOYMENT_STANDARD'].astype(str) + therap_prod['EMPLOYMENT_PERCENTILE'].astype(str) + therap_prod['HEALTH_SAFETY_RAW'].astype(str) + therap_prod['HEALTH_SAFETY_STANDARD'].astype(str) + therap_prod['HEALTH_SAFETY_PERCENTILE'].astype(str) + therap_prod['SOCIAL_RAW'].astype(str) + therap_prod['SOCIAL_STANDARD'].astype(str) + therap_prod['SOCIAL_PERCENTILE'].astype(str) + therap_prod['TOTAL_SCORE_NEEDED_INDEX'].astype(str) + therap_prod['EXCEPTION_MEDICAL_TOTAL'].astype(str) + therap_prod['EXCEPTION_BEHAVE_TOTAL'].astype(str)
sorted_therap = therap_prod.sort_values(['SIS_ID'], ascending=True)

sis_online['Tracking Number'] = sis_online['Tracking Number'].fillna(0)
sis_online['Concatenated_SIS_Online'] = sis_online['SIS_ID'].astype(str) + sis_online['First_Name'] + sis_online['Last_Name'] + sis_online['Tracking Number'].astype(str) + sis_online['HOME_LIVING_RAW'].astype(str) + sis_online['HOME_LIVING_STANDARD'].astype(str) + sis_online['HOME_LIVING_PERCENTILE'].astype(str) + sis_online['COMMUNITY_LIVING_RAW'].astype(str) + sis_online['COMMUNITY_LIVING_STANDARD'].astype(str) + sis_online['COMMUNITY_LIVING_PERCENTILE'].astype(str) + sis_online['LIFELONG_LEARNING_RAW'].astype(str) + sis_online['LIFELONG_LEARNING_STANDARD'].astype(str) + sis_online['LIFELONG_LEARNING_PERCENTILE'].astype(str) + sis_online['EMPLOYMENT_RAW'].astype(str) + sis_online['EMPLOYMENT_STANDARD'].astype(str) + sis_online['EMPLOYMENT_PERCENTILE'].astype(str) + sis_online['HEALTH_SAFETY_RAW'].astype(str) + sis_online['HEALTH_SAFETY_STANDARD'].astype(str) + sis_online['HEALTH_SAFETY_PERCENTILE'].astype(str) + sis_online['SOCIAL_RAW'].astype(str) + sis_online['SOCIAL_STANDARD'].astype(str) + sis_online['SOCIAL_PERCENTILE'].astype(str) + sis_online['TOTAL_SCORE_NEEDED_INDEX'].astype(str) + sis_online['EXCEPTION_MEDICAL_TOTAL'].astype(str) + sis_online['EXCEPTION_BEHAVE_TOTAL'].astype(str)
#.str[:-2] #have to be generic
sorted_sis_online = sis_online.sort_values(['SIS_ID'], ascending=True)
#print(therap_prod.info())


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_therap = pd.ExcelWriter(output_therap_path, engine='xlsxwriter')
writer_sis_online = pd.ExcelWriter(output_sis_online_path, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
sorted_therap.to_excel(writer_therap, index=False, sheet_name='Sheet1')
sorted_sis_online.to_excel(writer_sis_online, index=False, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer_therap.save()
writer_sis_online.save()


print('Successfull!')