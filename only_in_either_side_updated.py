#Sequence:5 updated
import pandas as pd
import numpy as np

#r_input_file_with_date_fields = 'r_prod_with_all_date_fields_8_aug_2017_new(2).xlsx'
date = '8_sep_2017'

input_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_file_all_date_fields_" + date + ".xlsx"
input_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\selected_columns_r_online_all_date_fields_" + date + ".xlsx"
output_only_at_r_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\exists_only_at_r_" + date + ".xlsx"
output_only_at_r_online_path = r"C:\Users\rasha\PycharmProjects\r_Report_Helper\output\exists_only_at_r_online_" + date + ".xlsx"


input_r_dataframe = pd.read_excel(input_r_path, sheetname='Sheet1', na_values=['NA'])
input_r_online_dataframe = pd.read_excel(input_r_online_path, sheetname='Sheet1', na_values=['NA'])

################################################## exists_only_at_r ################################################

r_only_dataframe = pd.merge(input_r_dataframe, input_r_online_dataframe, on=['r_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
r_only_dataframe['Exist'] = np.where(r_only_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (r_only_dataframe[['r_ID', 'Tracking Number', 'Exist', 'dateUpdated_r_Online']])
#print(df['Exist'].unique())
print(r_only_dataframe['Exist'].value_counts())


r_only = r_only_dataframe['Exist'] == False
#print(r_only_dataframe[r_only][:5])
print(r_only_dataframe[r_only]['Exist'].value_counts())
exists_only_at_r = r_only_dataframe[r_only][['r_ID', 'r_TRACKING_NUM', 'r_STATUS', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE']]
#print(exists_only_at_r[:5])

exists_only_at_r["CREATED"] = exists_only_at_r["CREATED"].map(lambda x: str(x)[:10])
exists_only_at_r["UPDATED"] = exists_only_at_r["UPDATED"].map(lambda x: str(x)[:10])
exists_only_at_r["STATUS_CHANGE_DATE"] = exists_only_at_r["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:10])
exists_only_at_r["COMPLETED_DATE"] = exists_only_at_r["COMPLETED_DATE"].map(lambda x: str(x)[:10])

writer_only_at_r = pd.ExcelWriter(output_only_at_r_path, engine='xlsxwriter')
exists_only_at_r.to_excel(writer_only_at_r, index=False, sheet_name='Sheet1')
writer_only_at_r.save()

print('Exist Only at r Successfull!')

############################################## exists_only_at_r_online ######################################################

r_only_dataframe = pd.merge(input_r_dataframe, input_r_online_dataframe, on=['r_ID'], how='right', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
r_only_dataframe['Exist'] = np.where(r_only_dataframe.Exist == 'both', True, False)
#print (df[:10])
print (r_only_dataframe[['r_ID', 'Tracking Number', 'statusText', 'locked', 'dateUpdated_r_Online', 'Exist']][:5])
#print(df['Exist'].unique())
print(r_only_dataframe['Exist'].value_counts())

######################################## r Online r_STATUS ###########################################
def r_online_r_status_formation(r_only_dataframe):
    if r_only_dataframe['locked'] == True:
        r_only_dataframe['r_STATUS'] = r_only_dataframe['statusText'] + '_LOCKED'
    else:
        r_only_dataframe['r_STATUS'] = r_only_dataframe['statusText']
    return r_only_dataframe['r_STATUS']

r_only_dataframe['r_STATUS'] = r_only_dataframe.apply(r_online_r_status_formation, axis=1)

is_locked = r_only_dataframe['locked'] == True
not_locked = r_only_dataframe['locked'] == False
# print(r_only_dataframe[is_locked]['locked'].value_counts())
# r_only_dataframe['r_STATUS'] = r_only_dataframe['statusText']
# r_only_dataframe[is_locked]['r_STATUS'] = r_only_dataframe[is_locked]['statusText'] + '_LOCKED'
#exists_only_at_r_online['r_STATUS'] = exists_only_at_r_online['statusText']
#exists_only_at_r_online['r_STATUS'] = exists_only_at_r_online['statusText'] + '_' + exists_only_at_r_online['locked'].astype(str)
#print(r_only_dataframe['r_STATUS'][:5])
print(r_only_dataframe[not_locked][['r_ID', 'Tracking Number', 'statusText', 'locked', 'r_STATUS', 'statusChangeDate', 'Completed Date', 'dateUpdated_r_Online']])

######################################## r Online r_STATUS ###########################################

r_only = r_only_dataframe['Exist'] == False
#is_locked = r_only_dataframe['locked'] == True
#print(r_only_dataframe[r_only][:5])
print(r_only_dataframe[r_only]['Exist'].value_counts())
#r_only_dataframe[r_only]['r_STATUS'] = r_only_dataframe[r_only].fillna('')['statusText'] + r_only_dataframe[r_only].fillna('')['locked'].astype(int).astype(str)
exists_only_at_r_online = r_only_dataframe[r_only][['r_ID', 'Tracking Number', 'r_STATUS', 'statusChangeDate', 'Completed Date', 'dateUpdated_r_Online']]
print(exists_only_at_r_online)

exists_only_at_r_online["statusChangeDate"] = exists_only_at_r_online["statusChangeDate"].str.split(' ').str[0]                 #.map(lambda x: str(x)[:10])
exists_only_at_r_online["dateUpdated_r_Online"] = exists_only_at_r_online["dateUpdated_r_Online"].str.split(' ').str[0]     #.map(lambda x: str(x)[:10])

exists_only_at_r_online["statusChangeDate"] = pd.to_datetime(exists_only_at_r_online["statusChangeDate"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
exists_only_at_r_online["Completed Date"] = pd.to_datetime(exists_only_at_r_online["Completed Date"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
exists_only_at_r_online["dateUpdated_r_Online"] = pd.to_datetime(exists_only_at_r_online["dateUpdated_r_Online"]).apply(lambda x: x.strftime('%m/%d/%Y')if not pd.isnull(x) else '')
# exists_only_at_r_online["statusChangeDate"] = exists_only_at_r_online["statusChangeDate"].dt.strftime("%d-%b-%y")
# exists_only_at_r_online["Completed Date"] = exists_only_at_r_online["Completed Date"].dt.strftime("%d-%b-%y")
# exists_only_at_r_online["dateUpdated_r_Online"] = exists_only_at_r_online["dateUpdated_r_Online"].dt.strftime("%d-%b-%y")

writer_only_at_r_online = pd.ExcelWriter(output_only_at_r_online_path, engine='xlsxwriter')
exists_only_at_r_online.to_excel(writer_only_at_r_online, index=False, sheet_name='Sheet1')
writer_only_at_r_online.save()

print('Exist Only at r Online Successfull!')
