# version 2, extract data directly from traveler sheet
# create one excel file per order number
# Fan Wu, Mar 28, 2021

import pandas as pd
import numpy as np
import sys
import os 

input_dir = os.path.join('..', 'Files2sort')
traveler_sheet_dir = os.path.join('..', 'TRAVELER SHEET-2021-UPTODATE_R1.1.xlsx')
traveler_sheet = pd.read_excel(traveler_sheet_dir, usecols = [1, 2, 3, 6, 9], header =0, skiprows=[1,2])
keys = pd.read_excel('Sort_keys.xlsx', index_col=0)
probe_convert_table = pd.read_excel('ProbeTypeConvert.xlsx', index_col=0)

def find_opens(data):
    opens = []
    opens = data[data['Impedance(MOhm)'] > thresh_open].index.tolist()
    return opens

def find_not_opens(data):
    not_opens = []
    not_opens = data[data['Impedance(MOhm)'] <= thresh_open].index.tolist()
    return not_opens

def find_shorts(data):
    shorts = []
    not_opens = find_not_opens(data)
    mean = sum(data['Impedance(MOhm)'][not_opens])/len(not_opens)
    shorts = data[data['Impedance(MOhm)'] < mean*0.65].index.tolist()
    return shorts
    
def color_z(val):
    if val > thresh_open:
        color = 'red'
    elif val < thresh_short:
        color = 'blue'
    else:
        color = 'black'
    return 'color: %s' % color

def convert_probe(DBC_probe):
    if DBC_probe in probe_convert_table.index:
        return probe_convert_table.loc[DBC_probe]['CNT']
    else:
        print('Warning: this probe does not exist in the catalog')
        return DBC_probe

order_num = []
for fname in os.listdir(input_dir):
    if '.txt' in fname:       
        part_num = fname[:fname.find('.')]
        part_num = 'PN ' + part_num
        this_part = traveler_sheet.loc[traveler_sheet['P/N'] == part_num]
        order_num.append(this_part['ORDER #'].values[0])

if len(np.unique(order_num)) > 1:
    print('Error: there is more than one order contained in the Files2sort directory')
    sys.exit()
else:
    output_dir = 'Datasheets/Datasheet-' + order_num[0] + '.xlsx'

with pd.ExcelWriter(output_dir, engine='xlsxwriter') as writer:
    for fname in os.listdir(input_dir):
        if '.txt' in fname:
            part_num_raw = fname[:fname.find('.')]
            part_num = 'PN ' + part_num_raw
            this_part = traveler_sheet.loc[traveler_sheet['P/N'] == part_num]
            assy_type = this_part['ASSY #'].values[0]          
            probe_type = this_part['PROBE TYPE'].values[0]
            # Converting probe type from DBC to CNT conention
            probe_type = convert_probe(probe_type)

            if output_dir == '':
                output_dir = 'Datasheet-' + order_num + '.xlsx'
                
            path = os.path.join(input_dir, fname)
            impedance = pd.read_table(
                path, 
                engine ='python', 
                delim_whitespace=True, 
                skiprows=3, 
                skipfooter=1, 
                index_col=0, 
                names=['Impedance(MOhm)','Phase']
            )
            impedance['Channel']=keys[this_part['ASSY #']]
            impedance = impedance.dropna()
            impedance_sorted = impedance.sort_values('Channel')
            impedance_sorted = impedance_sorted.set_index('Channel')

            # opens and shorts definition subject to change
            thresh_open = 0.15
            not_opens = find_not_opens(impedance_sorted)
            mean = sum(impedance_sorted['Impedance(MOhm)'][not_opens])/len(not_opens)
            thresh_short = mean * 0.65
            num_opens = find_opens(impedance_sorted)
            num_shorts = find_shorts(impedance_sorted)

            impedance_sorted.style.\
                applymap(color_z, subset=['Impedance(MOhm)']).\
                to_excel(writer, sheet_name = part_num_raw)

            # Generate another table summarizing the impedance
            summary = pd.DataFrame([['Opens: ' + str(num_opens), 'Shorts: ' + str(num_shorts)]], columns=[assy_type, probe_type])
            summary.to_excel(writer, sheet_name= part_num_raw, startcol=4, index=False)



