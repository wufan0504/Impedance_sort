# version 3, extract data directly from traveler sheet
# create one excel file per order number
# can run multiple sales orders at a time
# can handle both .csv and .txt
# Fan Wu, Jul 12, 2023

import pandas as pd
import numpy as np
import sys
import os
import glob

input_dir = os.path.join('..', 'FilesACRO2sort')
traveler_sheet_dir = os.path.join('..', 'TRAVELER SHEET-2022-UPTODATE_R1.0.xlsx')
traveler_sheet = pd.read_excel(traveler_sheet_dir, usecols = [1, 2, 3, 6, 9, 11, 12], header =0, skiprows=[1,2])
# 1: date; 2: order#; 3: Part num; 6: Probe type; 9: Assy type; 11: Sharp; 12: Fiber
keys = pd.read_excel('Sort_keys.xlsx', index_col=0)

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


# get a list of all files in directory
files = glob.glob(os.path.join(input_dir, '*'))

if len(files) == 0:  # replace with your actual condition
    sys.exit("Error: no files exist in the directory.")
# sort files by name, this is important for the later for loop operation
files.sort()

# checking to see if there is only one Order number in the Files2Sort list
orders = []
for fname in os.listdir(input_dir):
    if fname == '.DS_Store':
        continue
    orders.append(fname[:fname.find('-')])

#this is the list of all orders in the directory, and determines the number of output datasheets
orders = np.unique(orders)
print('these are the orders, PLEASE CHECK')
print(orders)

for order in orders:
    output_dir = os.path.join('..', 'Datasheets_ACRO')
    #os.mkdir(output_dir)
      
    datasheet = os.path.join(output_dir, ('Datasheet-' + order + '.xlsx'))

    with pd.ExcelWriter(datasheet, engine='xlsxwriter') as writer:
        count = 0
        for fname in os.listdir(input_dir):
            # an inefficient loop that goes through every file each time it iterates through every sales order
            if fname[:fname.find('-')] == order:   
                print(fname)        
                part_num = fname[:fname.find('.')]
                this_part = traveler_sheet.loc[traveler_sheet['P/N'] == part_num]
                probe_type = this_part['PROBE TYPE'].values[0]
                sharp = this_part['SHARPEN'].values[0]
                if pd.isnull(sharp):
                    sharp = ''
                fiber = this_part['FIBER'].values[0]
                if pd.isnull(fiber):
                    fiber = ''
                path = os.path.join(input_dir, fname)

                if '.txt' in fname:
                    count = count + 1          
                    impedance_temp = pd.read_table(
                        path, 
                        engine ='python', 
                        delim_whitespace=True, 
                        skiprows=3, 
                        skipfooter=1, 
                        index_col=0, 
                        names=['Z(MOhm)','Phase'],
                        encoding = 'unicode_escape'
                    )
                    impedance_temp['Channel']=keys.loc[:, this_part['ASSY #']]
                    impedance_temp = impedance_temp.dropna()
                    impedance_temp = impedance_temp[['Channel', 'Z(MOhm)', 'Phase']]
                    # sorting channels
                    impedance = impedance_temp.sort_values('Channel')

                    #print(impedance['Z(MOhm)'])

                elif '.csv' in fname: 
                    count = count + 1 
                    impedance = pd.read_csv(path, usecols=[0, 4, 5])
                    impedance['Impedance Magnitude at 1000 Hz (ohms)'] = impedance['Impedance Magnitude at 1000 Hz (ohms)']/1000000
                    impedance.rename(columns={'Impedance Magnitude at 1000 Hz (ohms)': 'Z(MOhm)'}, inplace = True)
                    impedance.rename(columns={'Impedance Phase at 1000 Hz (degrees)': 'Phase'}, inplace = True)

                    #print(impedance['Z(MOhm)'])
                    
                else:
                    sys.exit('Error: check impedance file format')

                thresh_Agrade = 0.9 # accept probes with > 90% working channels
                thresh_open = 0.15
                not_opens = find_not_opens(impedance)
                mean = sum(impedance['Z(MOhm)'][not_opens])/len(not_opens)
                thresh_short = mean * 0.65
                opens = find_opens(impedance)
                shorts = find_shorts(impedance)
                impedance.loc[:, 'Comment'] = '-'
                impedance.loc[opens, 'Comment'] = 'Open'
                impedance.loc[shorts, 'Comment'] = 'Short'

                if (len(opens) + len(shorts)) / (len(opens) + len(not_opens)) > (1 - thresh_Agrade):
                    print("\033[91m {}\033[00m" .format(part_num + ' from order ' + order + ' has ' + str(len(opens)) + 
                          ' number of opens and ' + str(len(shorts)) + ' number of shorts.'))

                startcol = (count-1)*5
                impedance.style.set_properties(**{
                    'text-align': 'center',
                    'font-size': '9pt'
                    }
                ).to_excel(writer, startrow = 2, startcol = startcol, index = False, sheet_name='Impedance data')

                # Styling worksheet
                worksheet = writer.sheets['Impedance data']
                worksheet.write(0, startcol, 'PN ' + part_num)
                worksheet.write(1, startcol, probe_type)
                worksheet.write(1, startcol + 1, sharp)
                worksheet.write(1, startcol + 2, fiber)
                worksheet.set_page_view()
                worksheet.set_paper(9) #A4
                #worksheet.set_header(assy_type)
                worksheet.set_footer('Diagnostic Biochips')


print('All done!')

