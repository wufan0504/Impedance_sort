# Calculates yield over time for direct orders

import pandas as pd
import numpy as np
import sys
import os 
#import matplotlib.pyplot as plt

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

input_dir = os.path.dirname('/Users/fanwu/OneDrive - Diagnostic Biochips/Documents/Processing Team/3_POST-FAB/1_PACKAGING/1_IMP DATA/')

probe64_yield = pd.DataFrame()
probe128_yield = pd.DataFrame()
for fname in os.listdir(input_dir):
        if '.csv' in fname and len(fname) == 8:
            part_num = fname[:fname.find('.')]
            #print(part_num)
                
            path = os.path.join(input_dir, fname)
            impedance = pd.read_csv(path, usecols=[0, 4])
                        
            # opens and shorts definition subject to change
            thresh_open = 15e4
            not_opens = find_not_opens(impedance)
            if len(not_opens) > 0:
                mean = sum(impedance['Impedance Magnitude at 1000 Hz (ohms)'][not_opens])/len(not_opens)
                thresh_short = mean * 0.65
                shorts = find_shorts(impedance)
            else:
                mean = sum(impedance['Impedance Magnitude at 1000 Hz (ohms)'])/len(impedance['Impedance Magnitude at 1000 Hz (ohms)'])
                shorts = []
            
            opens = find_opens(impedance)

            num_opens = len(opens)
            num_shorts = len(shorts)
            new_probe_yield = {'ID':part_num, 'Opens': int(num_opens), "Shorts": num_shorts}
            if len(impedance) > 100:
                probe128_yield = probe128_yield.append(new_probe_yield, ignore_index=True)
            else:
                probe64_yield = probe64_yield.append(new_probe_yield, ignore_index=True)

#print(probe_yield)
#print(sum(probe_yield['Shorts']))

probe128_yield.to_excel("128ch_yield.xlsx")
probe64_yield.to_excel("64ch_yield.xlsx")

""" plt.plot(probe_yield['ID'], probe_yield['Opens'] + probe_yield['Shorts'], 'go-', linewidth=1.5, markersize=6, label = 'Bad chan total')
plt.plot(probe_yield['ID'], probe_yield['Opens'], 'ro--', linewidth=1, markersize=4, label = 'Opens')
plt.plot(probe_yield['ID'], probe_yield['Shorts'], 'bo--', linewidth=1, markersize=4,label = 'Shorts')
plt.legend()
plt.show() """