import pandas as pd
import numpy as np

df = pd.read_excel('ТПП.xlsx')
kalculation = pd.read_excel('Kalculation.xls')
list_of_sizes = pd.read_excel('Розміри.xlsx')

df.fillna(0, inplace=True)
kalculation.columns = kalculation.loc[0,:]
kalculation.drop(0, inplace = True)

col = df.columns
df.iloc[:, 4] = df.iloc[:, 4].astype(str)
df.iloc[:, 0] = df.iloc[:, 0].astype(str)
df['розмір'] = 0

df2 = pd.DataFrame()
for i in range(len(df)):
    df['розмір'][i] = str(df.loc[i, col[5]]) + 'x' + str(df.loc[i, col[6]]) + 'x' + str(df.loc[i, col[7]])

df2 = df[df['розмір'].isin(list_of_sizes['потрібно'])]

dff = pd.DataFrame()
dff['name'] = range(len(df))
dff['%'] = range(len(df))

for i in range(len(df2)):
    kalculation1 = kalculation.copy()
    df1 = df2.reset_index()

    kalculation1.кількість[2] = df1.loc[i, 'ПЕ']
    kalculation1.кількість[6] = df1['Wh 70'][i] + df1['Wh 80'][i] + df1['Br 70'][i] + df1['Br 80'][i]
    kalculation1.кількість[7] = df1['Br 70'][i] + df1['Br 80'][i]
    kalculation1.кількість[8] = df1['Wh 70'][i] + df1['Wh 80'][i]

    kalculation1.кількість[11] = df1['концентрат'][i]
    kalculation1.кількість[12] = df1['дисперсія'][i]
    kalculation1.кількість[13] = df1['пов-акт. зас.'][i]
    kalculation1.кількість[10] = kalculation1['кількість'][10:13].sum()
    kalculation1.кількість[14] = df1['клей'][i]

    kalculation1['Сума попередня, грн'][18] = df1['звв, грн'][i]
    kalculation1['Сума попередня, грн'][19] = df1['з/п, грн'][i]
    kalculation1['Сума попередня, грн'][20] = df1['опер. витр. грн'][i]
    kalculation1['Сума попередня, грн'][21] = df1['адмін. Витрати, грн'][i]
    kalculation1['Сума попередня, грн'][22] = df1['амортизація, грн'][i]
    kalculation1['Сума попередня, грн'][23] = kalculation1['Сума попередня, грн'][17:22].sum()

    # ==============================================================================================
    kalculation1['Сума попередня, грн'][2] = df1.loc[i, 'ПЕ, грн']
    kalculation1['Сума попередня, грн'][3] = df1.loc[i, 'ПЕ, грн']
    kalculation1['Сума попередня, грн'][6] = df1.loc[i, 'Wh, грн'] + df1.loc[i, 'Br, грн']
    kalculation1['Сума попередня, грн'][7] = df1.loc[i, 'Br, грн']
    kalculation1['Сума попередня, грн'][8] = df1.loc[i, 'Wh, грн']

    kalculation1['Сума попередня, грн'][11] = df1.loc[i, 'Концентрат, грн']
    kalculation1['Сума попередня, грн'][12] = df1.loc[i, 'дисп., грн']
    kalculation1['Сума попередня, грн'][13] = df1.loc[i, 'ПАЗ, грн']
    kalculation1['Сума попередня, грн'][10] = kalculation1['Сума попередня, грн'][10:14].sum()
    kalculation1['Сума попередня, грн'][14] = df1.loc[i, 'клей, грн']
    kalculation1['Сума попередня, грн'][15] = kalculation1.loc[10, 'Сума попередня, грн'] + kalculation1.loc[6, 'Сума попередня, грн'] + kalculation1.loc[14, 'Сума попередня, грн']

    kalculation1['Сума попередня, грн'][18] = df1.loc[i, 'звв, грн']
    kalculation1['Сума попередня, грн'][19] = df1.loc[i, 'з/п, грн']
    kalculation1['Сума попередня, грн'][20] = df1.loc[i, 'опер. витр. грн']
    kalculation1['Сума попередня, грн'][21] = df1.loc[i, 'адмін. Витрати, грн']
    kalculation1['Сума попередня, грн'][22] = df1.loc[i, 'амортизація, грн']
    kalculation1['Сума попередня, грн'][23] = kalculation1['Сума попередня, грн'][17:22].sum()

    kalculation1['Сума попередня, грн'][25] = kalculation1.loc[15, 'Сума попередня, грн'] + kalculation1.loc[3, 'Сума попередня, грн'] + kalculation1.loc[23, 'Сума попередня, грн']
    kalculation1['Сума попередня, грн'][27] = str((100*kalculation1.loc[15, 'Сума попередня, грн']/kalculation1.loc[25, 'Сума попередня, грн']).round(2)) + ' %'

    name = str('{}x{}x{}.xlsx'.format(df1[col[5]][i],df1[col[6]][i],df1[col[7]][i]))
    print(kalculation1)
    kalculation1.to_excel(name, index=False)