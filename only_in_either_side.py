#Sequence:5
import pandas as pd
import numpy as np

therap_input_file_with_date_fields = 'Therap_prod_with_all_date_fields_8_aug_2017_new(2).xlsx'
date = '8_aug_2017'

input_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\input" + "\\"  + therap_input_file_with_date_fields
input_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\selected_columns_sis_online_all_date_fields_" + date + ".xlsx"
output_only_at_therap_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_only_at_therap_" + date + ".xlsx"
output_only_at_sis_online_path = r"C:\Users\rasha\PycharmProjects\SIS_Report_Helper\output\exists_only_at_sis_online_" + date + ".xlsx"


input_therap_dataframe = pd.read_excel(input_therap_path, sheetname='Export Worksheet', na_values=['NA'])
input_sis_online_dataframe = pd.read_excel(input_sis_online_path, sheetname='Sheet1', na_values=['NA'])

################################################## exists_only_at_therap ################################################

therap_only_dataframe = pd.merge(input_therap_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='left', indicator='Exist')
#df.drop('Rating', inplace=True, axis=1)
therap_only_dataframe['Exist'] = np.where(therap_only_dataframe.Exist == 'both', True, False)
#print (df[:10])
    #print (therap_only_dataframe[['SIS_ID', 'Tracking Number', 'Exist', 'dateUpdated_SIS_Online']])
#print(df['Exist'].unique())
print(therap_only_dataframe['Exist'].value_counts())


therap_only = therap_only_dataframe['Exist'] == False
#print(therap_only_dataframe[therap_only][:5])
print(therap_only_dataframe[therap_only]['Exist'].value_counts())
exists_only_at_therap = therap_only_dataframe[therap_only][['SIS_ID', 'SIS_TRACKING_NUM', 'SIS_STATUS', 'CREATED', 'UPDATED', 'STATUS_CHANGE_DATE', 'COMPLETED_DATE']]
#print(exists_only_at_therap[:5])

exists_only_at_therap["CREATED"] = exists_only_at_therap["CREATED"].map(lambda x: str(x)[:10])
exists_only_at_therap["UPDATED"] = exists_only_at_therap["UPDATED"].map(lambda x: str(x)[:10])
exists_only_at_therap["STATUS_CHANGE_DATE"] = exists_only_at_therap["STATUS_CHANGE_DATE"].map(lambda x: str(x)[:10])
exists_only_at_therap["COMPLETED_DATE"] = exists_only_at_therap["COMPLETED_DATE"].map(lambda x: str(x)[:10])

writer_only_at_therap = pd.ExcelWriter(output_only_at_therap_path, engine='xlsxwriter')
exists_only_at_therap.to_excel(writer_only_at_therap, index=False, sheet_name='Sheet1')
writer_only_at_therap.save()

print('Exist Only at Therap Successfull!')

############################################## exists_only_at_sis_online ######################################################

sis_only_dataframe = pd.merge(input_therap_dataframe, input_sis_online_dataframe, on=['SIS_ID'], how='right', indicator='Exist')
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
