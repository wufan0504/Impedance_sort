import pandas as pd
import os 

directory = 'Files2sort'
keys = pd.read_excel('Sort_keys.xlsx', index_col=0)

with pd.ExcelWriter('Impedances.xlsx') as writer:
    for fname in os.listdir(directory):
        path = os.path.join(directory, fname)
        assy_start = fname.find('-')
        assy_end = fname.find('.')
        assy_type = fname[assy_start+1:assy_end]
        part_num = fname[0:assy_start]
    
        impedance = pd.read_table(
            path, 
            engine ='python', 
            delim_whitespace=True, 
            skiprows=3, 
            skipfooter=1, 
            index_col=0, 
            names=['Impedance(MOhm)','Phase']
        )
        impedance[assy_type + ' Map']=keys[assy_type]
        impedance_sorted = impedance.sort_values(assy_type + ' Map')
        impedance_sorted = impedance_sorted.set_index(assy_type + ' Map')
        impedance_sorted.to_excel(writer, sheet_name=part_num)
    