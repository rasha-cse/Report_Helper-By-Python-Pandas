#Sequence:2
import pandas as pd

r_input_file_for_Concatenate = 'r_for_Concatenate_8_aug_2017.xlsx'
date = '8_aug_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input" + "\\"  + r_input_file_for_Concatenate
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input\selected_columns_sis_online_for_concatenate_" + date + ".xlsx"
output_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_r_" + date + ".xlsx"
output_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\concatenated_sis_online_" + date + ".xlsx"

r_prod = pd.read_excel(input_r_path, sheetname='Export Worksheet', na_values=['NA'], converters={'SIS_TRACKING_NUM': lambda x: str(x)})
sis_online = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'], converters={'Tracking Number': lambda x: str(x)})

#r_prod['Concatenated_r'] = r_prod.apply(lambda x:'%s%s%s%s' % (x['SIS_ID'],x['FIRST_NAME'],x['LAST_NAME'],x['SIS_TRACKING_NUM']),axis=1)
r_prod['Concatenated_r'] = r_prod['SIS_ID'].astype(str) + r_prod['FIRST_NAME'] + r_prod['LAST_NAME'] + r_prod['SIS_TRACKING_NUM'].astype(str)
sorted_r = r_prod.sort_values(['SIS_ID'], ascending=True)

sis_online['Tracking Number'] = sis_online['Tracking Number'].fillna(0)
sis_online['Concatenated_SIS_Online'] = sis_online['SIS_ID'].astype(str) + sis_online['First_Name'] + sis_online['Last_Name'] + sis_online['Tracking Number'].astype(str)#.str[:-2] #have to be generic
sorted_sis_online = sis_online.sort_values(['SIS_ID'], ascending=True)
#print(r_prod.info())


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer_r = pd.ExcelWriter(output_r_path, engine='xlsxwriter')
writer_sis_online = pd.ExcelWriter(output_sis_online_path, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
sorted_r.to_excel(writer_r, index=False, sheet_name='Sheet1')
sorted_sis_online.to_excel(writer_sis_online, index=False, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer_r.save()
writer_sis_online.save()


print('Successfull!')
