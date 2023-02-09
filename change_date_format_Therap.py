import pandas as pd

input_unformatted_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input\SIS_INFO_FXRAC.xlsx"
output_formatted_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\Fomatted_Date_SIS_INFO_Therap_Prod.xlsx"
input_unformatted_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_only_at_sis_online.xlsx"
output_formatted_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\Fomatted_Date_SIS_Online.xlsx"

input_dataframe_therap = pd.read_excel(input_unformatted_therap_path, sheetname='Export Worksheet', na_values=['NA'])
input_dataframe_sis_online = pd.read_excel(input_unformatted_sis_online_path, sheetname='Sheet1', na_values=['NA'])
#print(input_dataframe_therap[:10])

################################# Change Date Format Therap ###############################################

input_dataframe_therap["CREATED"] = input_dataframe_therap["CREATED"].map(lambda x: str(x)[:9])
input_dataframe_therap["UPDATED"] = input_dataframe_therap["UPDATED"].map(lambda x: str(x)[:9])
input_dataframe_therap["STATUS_CHANGE_DATE"] = input_dataframe_therap["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:9])
input_dataframe_therap["COMPLETED_DATE"] = input_dataframe_therap["COMPLETED_DATE"].map(lambda x: str(x)[:9])

# print(input_dataframe_therap[:10])

writer_sis_therap = pd.ExcelWriter(output_formatted_therap_path, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
input_dataframe_therap.to_excel(writer_sis_therap, index=False, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer_sis_therap.save()


################################# Change Date Format SIS Online ###############################################
input_dataframe_sis_online["statusChangeDate"] = pd.to_datetime(input_dataframe_sis_online["statusChangeDate"])
input_dataframe_sis_online["Completed Date"] = pd.to_datetime(input_dataframe_sis_online["Completed Date"])
input_dataframe_sis_online["dateUpdated_SIS_Online"] = pd.to_datetime(input_dataframe_sis_online["dateUpdated_SIS_Online"])
input_dataframe_sis_online["statusChangeDate"] = input_dataframe_sis_online["statusChangeDate"].dt.strftime("%d-%b-%y")
input_dataframe_sis_online["Completed Date"] = input_dataframe_sis_online["Completed Date"].dt.strftime("%d-%b-%y")
input_dataframe_sis_online["dateUpdated_SIS_Online"] = input_dataframe_sis_online["dateUpdated_SIS_Online"].dt.strftime("%d-%b-%y")


writer_sis_online = pd.ExcelWriter(output_formatted_sis_online_path, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
input_dataframe_sis_online.to_excel(writer_sis_online, index=False, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer_sis_online.save()

print('Change Date Format Successfull!')