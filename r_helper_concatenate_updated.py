#Sequence:2 updated
import pandas as pd

#r_input_file_for_Concatenate = 'r_for_Concatenate_8_aug_2017.xlsx'
date = '8_sep_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_file_for_concatenate_" + date + ".xlsx"
input_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input\selected_columns_r_online_for_concatenate_" + date + ".xlsx"
output_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\concatenated_r_" + date + ".xlsx"
output_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\concatenated_r_online_" + date + ".xlsx"

r_prod = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'], converters={'r_TRACKING_NUM': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })
r_online = pd.read_excel(input_r_online_path, sheetname='Sheet1', na_values=['NA'], converters={'Tracking Number': lambda x: str(x), 'HOME_LIVING_RAW': lambda x: str(x), 'HOME_LIVING_STANDARD': lambda x: str(x), 'HOME_LIVING_PERCENTILE': lambda x: str(x), 'COMMUNITY_LIVING_RAW': lambda x: str(x), 'COMMUNITY_LIVING_STANDARD': lambda x: str(x), 'COMMUNITY_LIVING_PERCENTILE': lambda x: str(x), 'LIFELONG_LEARNING_RAW': lambda x: str(x), 'LIFELONG_LEARNING_STANDARD': lambda x: str(x), 'LIFELONG_LEARNING_PERCENTILE': lambda x: str(x), 'EMPLOYMENT_RAW': lambda x: str(x), 'EMPLOYMENT_STANDARD': lambda x: str(x), 'EMPLOYMENT_PERCENTILE': lambda x: str(x), 'HEALTH_SAFETY_RAW': lambda x: str(x), 'HEALTH_SAFETY_STANDARD': lambda x: str(x), 'HEALTH_SAFETY_PERCENTILE': lambda x: str(x), 'SOCIAL_RAW': lambda x: str(x), 'SOCIAL_STANDARD': lambda x: str(x), 'SOCIAL_PERCENTILE': lambda x: str(x), 'TOTAL_SCORE_NEEDED_INDEX': lambda x: str(x), 'EXCEPTION_MEDICAL_TOTAL': lambda x: str(x), 'EXCEPTION_BEHAVE_TOTAL': lambda x: str(x) })

#r_prod['Concatenated_r'] = r_prod.apply(lambda x:'%s%s%s%s' % (x['r_ID'],x['FIRST_NAME'],x['LAST_NAME'],x['r_TRACKING_NUM']),axis=1)
r_prod['Concatenated_r'] = r_prod['r_ID'].astype(str) + r_prod['FIRST_NAME'] + r_prod['LAST_NAME'] + r_prod['r_TRACKING_NUM'].astype(str) + r_prod['HOME_LIVING_RAW'].astype(str) + r_prod['HOME_LIVING_STANDARD'].astype(str) + r_prod['HOME_LIVING_PERCENTILE'].astype(str) + r_prod['COMMUNITY_LIVING_RAW'].astype(str) + r_prod['COMMUNITY_LIVING_STANDARD'].astype(str) + r_prod['COMMUNITY_LIVING_PERCENTILE'].astype(str) + r_prod['LIFELONG_LEARNING_RAW'].astype(str) + r_prod['LIFELONG_LEARNING_STANDARD'].astype(str) + r_prod['LIFELONG_LEARNING_PERCENTILE'].astype(str) + r_prod['EMPLOYMENT_RAW'].astype(str) + r_prod['EMPLOYMENT_STANDARD'].astype(str) + r_prod['EMPLOYMENT_PERCENTILE'].astype(str) + r_prod['HEALTH_SAFETY_RAW'].astype(str) + r_prod['HEALTH_SAFETY_STANDARD'].astype(str) + r_prod['HEALTH_SAFETY_PERCENTILE'].astype(str) + r_prod['SOCIAL_RAW'].astype(str) + r_prod['SOCIAL_STANDARD'].astype(str) + r_prod['SOCIAL_PERCENTILE'].astype(str) + r_prod['TOTAL_SCORE_NEEDED_INDEX'].astype(str) + r_prod['EXCEPTION_MEDICAL_TOTAL'].astype(str) + r_prod['EXCEPTION_BEHAVE_TOTAL'].astype(str)
sorted_r = r_prod.sort_values(['r_ID'], ascending=True)

r_online['Tracking Number'] = r_online['Tracking Number'].fillna(0)
r_online['Concatenated_r_Online'] = r_online['r_ID'].astype(str) + r_online['First_Name'] + r_online['Last_Name'] + r_online['Tracking Number'].astype(str) + r_online['HOME_LIVING_RAW'].astype(str) + r_online['HOME_LIVING_STANDARD'].astype(str) + r_online['HOME_LIVING_PERCENTILE'].astype(str) + r_online['COMMUNITY_LIVING_RAW'].astype(str) + r_online['COMMUNITY_LIVING_STANDARD'].astype(str) + r_online['COMMUNITY_LIVING_PERCENTILE'].astype(str) + r_online['LIFELONG_LEARNING_RAW'].astype(str) + r_online['LIFELONG_LEARNING_STANDARD'].astype(str) + r_online['LIFELONG_LEARNING_PERCENTILE'].astype(str) + r_online['EMPLOYMENT_RAW'].astype(str) + r_online['EMPLOYMENT_STANDARD'].astype(str) + r_online['EMPLOYMENT_PERCENTILE'].astype(str) + r_online['HEALTH_SAFETY_RAW'].astype(str) + r_online['HEALTH_SAFETY_STANDARD'].astype(str) + r_online['HEALTH_SAFETY_PERCENTILE'].astype(str) + r_online['SOCIAL_RAW'].astype(str) + r_online['SOCIAL_STANDARD'].astype(str) + r_online['SOCIAL_PERCENTILE'].astype(str) + r_online['TOTAL_SCORE_NEEDED_INDEX'].astype(str) + r_online['EXCEPTION_MEDICAL_TOTAL'].astype(str) + r_online['EXCEPTION_BEHAVE_TOTAL'].astype(str)
#.str[:-2] #have to be generic
sorted_r_online = r_online.sort_values(['r_ID'], ascending=True)
#print(r_prod.info())


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_r = pd.ExcelWriter(output_r_path, engine='xlsxwriter')
writer_r_online = pd.ExcelWriter(output_r_online_path, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
sorted_r.to_excel(writer_r, index=False, sheet_name='Sheet1')
sorted_r_online.to_excel(writer_r_online, index=False, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer_r.save()
writer_r_online.save()


print('Successfull!')
