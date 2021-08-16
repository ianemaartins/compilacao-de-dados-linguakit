import pandas as pd
import os

def tag_counter(file_name):
    row_name = os.path.splitext(file_name)[0]
    afn_id = row_name.strip('afn')
    if afn_id[0] == '0':
        afn_id = afn_id[1:]
    sheet = pd.read_excel(file_name, engine='openpyxl').iloc[:, [0,2]].groupby('Etiqueta').count().transpose()
    sheet.index = pd.Index([row_name])
    sheet['ID'] = int(afn_id)
    return(sheet)

master = pd.DataFrame()

for file_name in os.listdir():
    if file_name.endswith(".xlsx"):
        # Read the file and run the counting
        new_count = tag_counter(file_name)

        # Merge into the main dataframe
        master = master.append(new_count).fillna(0)


# Sort by ID (this is VERY optional)
master = master.sort_values(by='ID').drop(['ID'], axis = 1)

#Total counts
total_counts = master.sum().to_frame(name = 'Total').transpose()

# Appending total count
master = master.append(total_counts).astype(int).rename_axis('Table')

# Writeout
master.to_excel('Aggregator.xlsx')
    