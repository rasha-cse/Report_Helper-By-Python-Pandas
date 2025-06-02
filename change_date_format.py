import pandas as pd

input_unformatted_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\input\r_INFO_FXRAC.xlsx"
output_formatted_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\Fomatted_Date_r_INFO_r_Prod.xlsx"
input_unformatted_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\exists_only_at_r_online.xlsx"
output_formatted_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\Fomatted_Date_r_Online.xlsx"

input_dataframe_r = pd.read_excel(input_unformatted_r_path, sheetname='Export Worksheet', na_values=['NA'])
input_dataframe_r_online = pd.read_excel(input_unformatted_r_online_path, sheetname='Sheet1', na_values=['NA'])
#print(input_dataframe_r[:10])

################################# Change Date Format r ###############################################

input_dataframe_r["CREATED"] = input_dataframe_r["CREATED"].map(lambda x: str(x)[:9])
input_dataframe_r["UPDATED"] = input_dataframe_r["UPDATED"].map(lambda x: str(x)[:9])
input_dataframe_r["STATUS_CHANGE_DATE"] = input_dataframe_r["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:9])
input_dataframe_r["COMPLETED_DATE"] = input_dataframe_r["COMPLETED_DATE"].map(lambda x: str(x)[:9])

# print(input_dataframe_r[:10])

writer_r_r = pd.ExcelWriter(output_formatted_r_path, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
input_dataframe_r.to_excel(writer_r_r, index=False, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer_r_r.save()


################################# Change Date Format r Online ###############################################
input_dataframe_r_online["statusChangeDate"] = pd.to_datetime(input_dataframe_r_online["statusChangeDate"])
input_dataframe_r_online["Completed Date"] = pd.to_datetime(input_dataframe_r_online["Completed Date"])
input_dataframe_r_online["dateUpdated_r_Online"] = pd.to_datetime(input_dataframe_r_online["dateUpdated_r_Online"])
input_dataframe_r_online["statusChangeDate"] = input_dataframe_r_online["statusChangeDate"].dt.strftime("%d-%b-%y")
input_dataframe_r_online["Completed Date"] = input_dataframe_r_online["Completed Date"].dt.strftime("%d-%b-%y")
input_dataframe_r_online["dateUpdated_r_Online"] = input_dataframe_r_online["dateUpdated_r_Online"].dt.strftime("%d-%b-%y")


writer_r_online = pd.ExcelWriter(output_formatted_r_online_path, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
input_dataframe_r_online.to_excel(writer_r_online, index=False, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer_r_online.save()

print('Change Date Format Successfull!')
