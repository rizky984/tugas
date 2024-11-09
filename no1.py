import pandas as pd
import os

data = pd.read_excel('C:\py\dataset.xlsx')
print(data.head())


prodi_list = data['Prodi'].unique()


if not os.path.exists('Output'):
    os.makedirs('Output')


def clean_sheet_name(sheet_name):
    invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '_')  
    return sheet_name

for prodi in prodi_list:
    
    prodi_data = data[data['Prodi'] == prodi]

    
    mk_kelas_list = prodi_data[['Mata Kuliah', 'Kode MK']].drop_duplicates()

    
    with pd.ExcelWriter(f'Output/{prodi}.xlsx', engine='xlsxwriter') as writer:
        for _, row in mk_kelas_list.iterrows():
            
            mk_data = prodi_data[(prodi_data['Mata Kuliah'] == row['Mata Kuliah']) &
                                 (prodi_data['Kode MK'] == row['Kode MK'])]

            
            sheet_name = f"{row['Mata Kuliah']}_{row['Kode MK']}"[:31]  

            
            sheet_name = clean_sheet_name(sheet_name)

            
            mk_data.to_excel(writer, sheet_name=sheet_name, index=False)