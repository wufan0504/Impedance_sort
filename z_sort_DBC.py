# version 2, extract data directly from traveler sheet
# create one excel file per order number
# Fan Wu, Mar 28, 2021

import pandas as pd
import numpy as np
import sys
import os 

input_dir = os.path.join('..', 'FilesDBC2sort')
output_dir = 'Datasheets/Datasheet.xlsx'

def find_opens(data):
    opens = []
    opens = data[data['Impedance Magnitude at 1000 Hz (ohms)'] > thresh_open].index.tolist()
    return opens

def find_not_opens(data):
    not_opens = []
    not_opens = data[data['Impedance Magnitude at 1000 Hz (ohms)'] <= thresh_open].index.tolist()
    return not_opens

def find_shorts(data):
    shorts = []
    not_opens = find_not_opens(data)
    mean = sum(data['Impedance Magnitude at 1000 Hz (ohms)'][not_opens])/len(not_opens)
    shorts = data[data['Impedance Magnitude at 1000 Hz (ohms)'] < mean*0.65].index.tolist()
    return shorts
    

with pd.ExcelWriter(output_dir, engine='xlsxwriter') as writer:
    count = 0
    for fname in os.listdir(input_dir):
        if '.csv' in fname:
            count = count + 1
            part_num = fname[:fname.find('.')]
            # need to label with part_num    
            path = os.path.join(input_dir, fname)
            impedance = pd.read_csv(path, usecols=[0, 4])
            
            # opens and shorts definition subject to change
            thresh_open = 15e4
            not_opens = find_not_opens(impedance)
            mean = sum(impedance['Impedance Magnitude at 1000 Hz (ohms)'][not_opens])/len(not_opens)
            thresh_short = mean * 0.65
            opens = find_opens(impedance)
            shorts = find_shorts(impedance)
            impedance.loc[:, 'Comment'] = '-'
            impedance.loc[opens, 'Comment'] = 'Open'
            impedance.loc[shorts, 'Comment'] = 'Short'

            impedance['Impedance Magnitude at 1000 Hz (ohms)'] = impedance['Impedance Magnitude at 1000 Hz (ohms)']/1000000
            impedance.rename(columns={"Impedance Magnitude at 1000 Hz (ohms)": "Impedance Mag at 1kHz (Mohm)"}, inplace = True)


            # sorting impedance
            # impedance_sorted = impedance.sort_values('Channel')
            #impedance_sorted = impedance_sorted.set_index('Channel')

            #write to excel
            # impedance_sorted.style.\
            #     applymap(color_z, subset=['Z(MOhm)']).\
            #     to_excel(writer, startrow = 2, sheet_name = part_num_raw)
            
            #style dataframe

            impedance.style.set_properties(**{
                    'text-align': 'center',
                    'font-size': '9pt'
                }
            ).to_excel(writer, startcol = (count-1)*4, index = False, sheet_name='Impedance data')


            # Generate another table summarizing the impedance
            # summary = pd.DataFrame([['Opens: ' + str(opens), 'Shorts: ' + str(shorts)]], columns=[assy_type, probe_type])
            # summary.to_excel(writer, sheet_name= part_num_raw, index=False)

# print('Assemblies in this datasheet: ', assy_dict)

