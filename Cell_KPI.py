
from functools import reduce

import numpy as np
import pandas as pd
import xlsxwriter


#import matplotlib.pyplot as plt

#def worstcount

df = pd.read_csv("KPI.csv")

#kpi=["ERAB_SETUP_SR","CSFB_PREP_SR","RRC_SETUP_SR","END_TO_END_EST_RATE"]

cells_all = df[["ENODEB_NAME","CELL_NAME"]].drop_duplicates()

df["FRAGMENT_DATE"] = pd.to_datetime(df["FRAGMENT_DATE"])

df96_erab = df[(df.ERAB_SETUP_SR<96) & (df.ERAB_SETUP_SR>0)].groupby(['CELL_NAME'])[['ERAB_SETUP_SR']].agg('count').reset_index()
df98_erab = df[(df.ERAB_SETUP_SR<98) & (df.ERAB_SETUP_SR>=96)].groupby(['CELL_NAME'])[['ERAB_SETUP_SR']].agg('count').reset_index()

df96_csfb = df[(df.CSFB_PREP_SR<96) & (df.CSFB_PREP_SR>0)].groupby(['CELL_NAME'])[['CSFB_PREP_SR']].agg('count').reset_index()
df98_csfb = df[(df.CSFB_PREP_SR<98) & (df.CSFB_PREP_SR>=96)].groupby(['CELL_NAME'])[['CSFB_PREP_SR']].agg('count').reset_index()

df96_rrc = df[(df.RRC_SETUP_SR<96) & (df.RRC_SETUP_SR>0)].groupby(['CELL_NAME'])[['RRC_SETUP_SR']].agg('count').reset_index()
df98_rrc = df[(df.RRC_SETUP_SR<98) & (df.RRC_SETUP_SR>=96)].groupby(['CELL_NAME'])[['RRC_SETUP_SR']].agg('count').reset_index()

df96_e2e = df[(df.END_TO_END_EST_RATE<96) & (df.END_TO_END_EST_RATE>0)].groupby(['CELL_NAME'])[['END_TO_END_EST_RATE']].agg('count').reset_index()
df98_e2e = df[(df.END_TO_END_EST_RATE<98) & (df.END_TO_END_EST_RATE>=96)].groupby(['CELL_NAME'])[['END_TO_END_EST_RATE']].agg('count').reset_index()

df3_drop = df[(df.DCR_VOLTE_V5R<100) & (df.DCR_VOLTE_V5R>=3)].groupby(['CELL_NAME'])[['DCR_VOLTE_V5R']].agg('count').reset_index()

worstcell = reduce(lambda left, right: pd.merge(left, right, on='CELL_NAME', how='left'), [cells_all,df96_erab,df96_csfb,df96_rrc,df96_e2e,df98_erab,df98_csfb,df98_rrc,df98_e2e,df3_drop])
worstcell.fillna(0, inplace=True)
worstcell.rename(
    columns={'CELL_NAME': 'CELL_NAME', 'ERAB_SETUP_SR_x': '# of ERAB<%96', 'CSFB_PREP_SR_x': '# of CSFB<%96' , 'RRC_SETUP_SR_x': '# of RRC<%96', 'END_TO_END_EST_RATE_x': '# of E2E<%96' ,'ERAB_SETUP_SR_y': '# of ERAB<%98', 'CSFB_PREP_SR_y': '# of CSFB<%98' , 'RRC_SETUP_SR_y': '# of RRC<%98', 'END_TO_END_EST_RATE_y': '# of E2E<%98', 'DCR_VOLTE_V5R_x': '# of Drop>%3'} ,
    inplace=True)

worstcell["TOTAL96"] = worstcell["# of ERAB<%96"] + worstcell["# of CSFB<%96"] + worstcell["# of RRC<%96"] + worstcell["# of E2E<%96"]
worstcell["TOTAL98"] = worstcell["# of ERAB<%98"] + worstcell["# of CSFB<%98"] + worstcell["# of RRC<%98"] + worstcell["# of E2E<%98"] + worstcell["DCR_VOLTE_V5R"]
worstcell["TOTAL"] = worstcell["TOTAL96"] + worstcell["TOTAL98"]
#worstcell_filtered = worstcell[(worstcell.TOTAL>0)]

print("Loading...")
worstcell_filter = worstcell[["CELL_NAME","TOTAL"]]
pivot_df = pd.merge(df,worstcell_filter,on="CELL_NAME",how="left")
pivot_df.fillna(0, inplace=True)
pivot_df = pivot_df[(pivot_df.TOTAL>0)]
print("Loading...")
#pivot_df.to_csv("pivot.csv")


print("Loading...")
table_erab = pd.pivot_table(pivot_df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["ERAB_SETUP_SR"],aggfunc=np.mean)
table_drop = pd.pivot_table(pivot_df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["DCR_VOLTE_V5R"],aggfunc=np.mean)
table_csfb = pd.pivot_table(pivot_df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["CSFB_PREP_SR"],aggfunc=np.mean)
table_rrc = pd.pivot_table(pivot_df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["RRC_SETUP_SR"],aggfunc=np.mean)
table_e2e = pd.pivot_table(pivot_df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["END_TO_END_EST_RATE"],aggfunc=np.mean)
#table_erab.ffill(inplace=True)

worstcell.to_csv("worstcell.csv")
table_erab.to_csv("erab.csv")
table_csfb.to_csv("csfb.csv")
table_rrc.to_csv("rrc.csv")
table_e2e.to_csv("e2e.csv")
table_drop.to_csv("volte_drop.csv")

print("Loading...")

with pd.ExcelWriter('KPI_control.xlsx') as writer:
    table_erab.to_excel(writer, sheet_name='ERAB_SR')
    table_rrc.to_excel(writer, sheet_name='RRC_SR')
    table_csfb.to_excel(writer, sheet_name='CSFB_SR')
    table_e2e.to_excel(writer, sheet_name='END_TO_END_SR')
    table_drop.to_excel(writer, sheet_name='VOLTE_DROP')
    worstcell.to_excel(writer, sheet_name='Worstcell')

print("Good jobs")



'''
#excel sheet lere yazdırma
writer = pd.ExcelWriter('Worst_Cell_List.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
worstcell.to_excel(writer, sheet_name='worst')
table_erab.to_excel(writer, sheet_name='erab_pivot')
#df3.to_excel(writer, sheet_name='Sheet3')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
'''



'''
print(df96)
df96.to_csv("df96.csv")
#print(df.head())

table = pd.pivot_table(df,index=["ENODEB_NAME","CELL_NAME"],columns=["FRAGMENT_DATE"],values=["ERAB_SETUP_SR"],aggfunc=np.mean)

#table["#of<96"] = table["2/10/21 12:00 AM"].apply(worstcount)
pd.set_option('display.max_columns',len(table))
print(table.columns.values.tolist())
print(table[('ERAB_SETUP_SR', '2/10/21 12:00 AM')])
table[]
'''

#print(table.to_string())