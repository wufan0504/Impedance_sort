# version 2, extract data directly from traveler sheet
# create one excel file per order number
# Fan Wu, Mar 28, 2021

import pandas as pd
import numpy as np
import sys
import os 

input_dir = os.path.join('..', 'FilesCNT2sort')
traveler_sheet_dir = os.path.join('..', 'TRAVELER SHEET-2022-UPTODATE_R1.0.xlsx')
traveler_sheet = pd.read_excel(traveler_sheet_dir, usecols = [1, 2, 3, 6, 9, 11, 12], header =0, skiprows=[1,2])
# 1: date; 2: order#; 3: Part num; 6: Probe type; 9: Assy type; 11: Sharp; 12: Fiber
keys = pd.read_excel('Sort_keys.xlsx', index_col=0)
probe_convert_table = pd.read_excel('ProbeTypeConvert.xlsx', index_col=0)

def find_opens(data):
    opens = []
    opens = data[data['Z(MOhm)'] > thresh_open].index.tolist()
    return opens

def find_not_opens(data):
    not_opens = []
    not_opens = data[data['Z(MOhm)'] <= thresh_open].index.tolist()
    return not_opens

def find_shorts(data):
    shorts = []
    not_opens = find_not_opens(data)
    mean = sum(data['Z(MOhm)'][not_opens])/len(not_opens)
    shorts = data[data['Z(MOhm)'] < mean*0.65].index.tolist()
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

# checking to see if there is only one Order number in the Files2Sort list
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

    output_dir = 'Datasheets/Datasheet-' + str(order_num[0]) + '.xlsx'

with pd.ExcelWriter(output_dir, engine='xlsxwriter') as writer:
    assy_dict = {}
    for fname in os.listdir(input_dir):
        if '.txt' in fname:
            part_num_raw = fname[:fname.find('.')]
            part_num = 'PN ' + part_num_raw
            this_part = traveler_sheet.loc[traveler_sheet['P/N'] == part_num]
            assy_type = this_part['ASSY #'].values[0]     
            if assy_type in assy_dict:
                assy_dict.update({assy_type: assy_dict[assy_type] + 1})
            else:
                assy_dict.update({assy_type: 1})
            
            sharp = this_part['SHARPEN'].values[0]
            if pd.isnull(sharp):
                sharp = ''
            fiber = this_part['FIBER'].values[0]
            if pd.isnull(fiber):
                fiber = ''
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
                names=['Z(MOhm)','Phase'],
            )
            impedance['Channel']=keys.loc[:, this_part['ASSY #']]
            impedance = impedance.dropna()
            impedance = impedance[['Channel', 'Z(MOhm)', 'Phase']]

            # opens and shorts definition subject to change
            thresh_open = 0.15
            not_opens = find_not_opens(impedance)
            mean = sum(impedance['Z(MOhm)'][not_opens])/len(not_opens)
            thresh_short = mean * 0.65
            opens = find_opens(impedance)
            shorts = find_shorts(impedance)
            impedance.loc[:, 'Comment'] = '-'
            impedance.loc[opens, 'Comment'] = 'Open'
            impedance.loc[shorts, 'Comment'] = 'Short'

            # sorting impedance
            impedance_sorted = impedance.sort_values('Channel')
            #impedance_sorted = impedance_sorted.set_index('Channel')

            #write to excel
            # impedance_sorted.style.\
            #     applymap(color_z, subset=['Z(MOhm)']).\
            #     to_excel(writer, startrow = 2, sheet_name = part_num_raw)
            
            #style dataframe
            if assy_dict[assy_type] % 2 == 1:
                startcol = assy_dict[assy_type] // 2 * 9
            else:
                startcol = (assy_dict[assy_type] // 2 -1) * 9 + 5

            impedance_sorted.style.set_properties(**{
                    'text-align': 'center',
                    'font-size': '9pt'
                }
            ).to_excel(writer, startrow = 2, startcol = startcol, index = False, sheet_name = assy_type)

            # Styling worksheet
            worksheet = writer.sheets[assy_type]
            worksheet.write(0, startcol, part_num)
            worksheet.write(1, startcol, probe_type)
            worksheet.write(1, startcol + 1, sharp)
            worksheet.write(1, startcol + 2, fiber)
            worksheet.set_page_view()
            worksheet.set_paper(9) #A4
            worksheet.set_header(assy_type)
            worksheet.set_footer('Cambridge NeuroTech')
            # worksheet.fit_to_pages(1,0)

            # Generate another table summarizing the impedance
            # summary = pd.DataFrame([['Opens: ' + str(opens), 'Shorts: ' + str(shorts)]], columns=[assy_type, probe_type])
            # summary.to_excel(writer, sheet_name= part_num_raw, index=False)

print('Assemblies in this datasheet: ', assy_dict)

