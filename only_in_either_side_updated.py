#Sequence:5 updated
import pandas as pd
import numpy as np

#r_input_file_with_date_fields = 'r_prod_with_all_date_fields_8_aug_2017_new(2).xlsx'
date = '8_sep_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_r_file_all_date_fields_" + date + ".xlsx"
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_sis_online_all_date_fields_" + date + ".xlsx"
output_only_at_r_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_only_at_r_" + date + ".xlsx"
output_only_at_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_only_at_sis_online_" + date + ".xlsx"


input_r_dataframe = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'])
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'])

################################################## exists_only_at_r ################################################

r_only_dataframe = pd.merge(input_r_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
r_only_dataframe['Exist'] = np.where(r_only_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (r_only_dataframe[['SIS_ID', 'Tracking Number', 'Exist', 'dateUpdated_SIS_Online']])
#print(df['Exist'].unique())
print(r_only_dataframe['Exist'].value_counts())


r_only = r_only_dataframe['Exist'] == False
#print(r_only_dataframe[r_only][:5])
print(r_only_dataframe[r_only]['Exist'].value_counts())
exists_only_at_r = r_only_dataframe[r_only][['SIS_ID', 'SIS_TRACKING_NUM', 'SIS_STATUS', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE']]
#print(exists_only_at_r[:5])

exists_only_at_r["CREATED"] = exists_only_at_r["CREATED"].map(lambda x: str(x)[:10])
exists_only_at_r["UPDATED"] = exists_only_at_r["UPDATED"].map(lambda x: str(x)[:10])
exists_only_at_r["STATUS_CHANGE_DATE"] = exists_only_at_r["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:10])
exists_only_at_r["COMPLETED_DATE"] = exists_only_at_r["COMPLETED_DATE"].map(lambda x: str(x)[:10])

writer_only_at_r = pd.ExcelWriter(output_only_at_r_path, engine='xlsxwriter')
exists_only_at_r.to_excel(writer_only_at_r, index=False, sheet_name='Sheet1')
writer_only_at_r.save()

print('Exist Only at r Successfull!')

############################################## exists_only_at_sis_online ######################################################

sis_only_dataframe = pd.merge(input_r_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='right', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
sis_only_dataframe['Exist'] = np.where(sis_only_dataframe.Exist == 'both', True, False)
#print (df[:10])
print (sis_only_dataframe[['SIS_ID', 'Tracking Number', 'statusText', 'locked', 'dateUpdated_SIS_Online', 'Exist']][:5])
#print(df['Exist'].unique())
print(sis_only_dataframe['Exist'].value_counts())

######################################## SIS Online SIS_STATUS ###########################################
def sis_online_sis_status_formation(sis_only_dataframe):
    if sis_only_dataframe['locked'] == True:
        sis_only_dataframe['SIS_STATUS'] = sis_only_dataframe['statusText'] + '_LOCKED'
    else:
        sis_only_dataframe['SIS_STATUS'] = sis_only_dataframe['statusText']
    return sis_only_dataframe['SIS_STATUS']

sis_only_dataframe['SIS_STATUS'] = sis_only_dataframe.apply(sis_online_sis_status_formation, axis=1)

is_locked = sis_only_dataframe['locked'] == True
not_locked = sis_only_dataframe['locked'] == False
# print(sis_only_dataframe[is_locked]['locked'].value_counts())
# sis_only_dataframe['SIS_STATUS'] = sis_only_dataframe['statusText']
# sis_only_dataframe[is_locked]['SIS_STATUS'] = sis_only_dataframe[is_locked]['statusText'] + '_LOCKED'
#exists_only_at_sis_online['SIS_STATUS'] = exists_only_at_sis_online['statusText']
#exists_only_at_sis_online['SIS_STATUS'] = exists_only_at_sis_online['statusText'] + '_' + exists_only_at_sis_online['locked'].astype(str)
#print(sis_only_dataframe['SIS_STATUS'][:5])
print(sis_only_dataframe[not_locked][['SIS_ID', 'Tracking Number', 'statusText', 'locked', 'SIS_STATUS', 'statusChangeDate', 'Completed Date', 'dateUpdated_SIS_Online']])

######################################## SIS Online SIS_STATUS ###########################################

sis_only = sis_only_dataframe['Exist'] == False
#is_locked = sis_only_dataframe['locked'] == True
#print(sis_only_dataframe[sis_only][:5])
print(sis_only_dataframe[sis_only]['Exist'].value_counts())
#sis_only_dataframe[sis_only]['SIS_STATUS'] = sis_only_dataframe[sis_only].fillna('')['statusText'] + sis_only_dataframe[sis_only].fillna('')['locked'].astype(int).astype(str)
exists_only_at_sis_online = sis_only_dataframe[sis_only][['SIS_ID', 'Tracking Number', 'SIS_STATUS', 'statusChangeDate', 'Completed Date', 'dateUpdated_SIS_Online']]
print(exists_only_at_sis_online)

exists_only_at_sis_online["statusChangeDate"] = exists_only_at_sis_online["statusChangeDate"].str.split(' ').str[0]                 #.map(lambda x: str(x)[:10])
exists_only_at_sis_online["dateUpdated_SIS_Online"] = exists_only_at_sis_online["dateUpdated_SIS_Online"].str.split(' ').str[0]     #.map(lambda x: str(x)[:10])

exists_only_at_sis_online["statusChangeDate"] = pd.to_datetime(exists_only_at_sis_online["statusChangeDate"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
exists_only_at_sis_online["Completed Date"] = pd.to_datetime(exists_only_at_sis_online["Completed Date"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
exists_only_at_sis_online["dateUpdated_SIS_Online"] = pd.to_datetime(exists_only_at_sis_online["dateUpdated_SIS_Online"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
# exists_only_at_sis_online["statusChangeDate"] = exists_only_at_sis_online["statusChangeDate"].dt.strftime("%d-%b-%y")
# exists_only_at_sis_online["Completed Date"] = exists_only_at_sis_online["Completed Date"].dt.strftime("%d-%b-%y")
# exists_only_at_sis_online["dateUpdated_SIS_Online"] = exists_only_at_sis_online["dateUpdated_SIS_Online"].dt.strftime("%d-%b-%y")

writer_only_at_sis_online = pd.ExcelWriter(output_only_at_sis_online_path, engine='xlsxwriter')
exists_only_at_sis_online.to_excel(writer_only_at_sis_online, index=False, sheet_name='Sheet1')
writer_only_at_sis_online.save()

print('Exist Only at SIS Online Successfull!')
