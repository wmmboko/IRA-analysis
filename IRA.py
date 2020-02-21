import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
https: // github.com/wmmboko/IRA-analysis

# Select your link
excel_file = '2018_Q4_Statistics-Locked.xlsx'

# %%
"""
This section identifies the sheet names to be used in analysis.
It then classifies the sheet as either lifeor general
"""
# ____________________________________________________________________________#

# Specifies which worksheet to read. I had to use the Table of Contents since
# naming was not consistent sheet to sheet

sheet_names = pd.read_excel(excel_file, sheet_name='Table of Contents',
                            header=6, usecols=[1, 2])
sheet_names['Link'] = sheet_names['Link'].apply(lambda x: x.strip("''"))
# identifies different report types
gen_sheets = sheet_names[sheet_names['Description'].str.contains('GENERAL')]
life_sheets = sheet_names[~sheet_names['Description'].str.contains('GENERAL')]

bal_sheets = sheet_names[sheet_names['Description'].str.contains('BALANCE')]
pnl_sheets = sheet_names[~sheet_names['Description'].str.contains('BALANCE')]


def ReadData(sheet):
    df = pd.read_excel(excel_file, sheet_name=sheet)
    return df


# identify general companies
ds1 = ReadData(gen_sheets['Link'][0])
rem_list = ['TOTAL', 'REINSURERS',
            'GRAND TOTAL', 'Amounts in Thousand Shillings']
gen_companies = list(ds1.iloc[4:, 1])
gen_companies_r = list(filter(lambda x: 'REINS' in x, gen_companies))
rem_list.extend(gen_companies_r)
gen_companies = list(filter(lambda x: x not in rem_list, gen_companies))

# identify life companies
ds1 = ReadData(life_sheets['Link'][1])
rem_list = ['TOTAL', 'REINSURERS',
            'GRAND TOTAL', 'Amounts in Thousand Shillings']
life_companies = list(ds1.iloc[4:, 1])
life_companies_r = list(filter(lambda x: 'REINS' in x, life_companies))
rem_list.extend(life_companies_r)
life_companies = list(filter(lambda x: x not in rem_list, life_companies))
del ds1
# premium summaries

gb_prem = ReadData("APPENDIX 13")
header = gb_prem[gb_prem.iloc[:, 1] == 'Company'].dropna(axis=1)
header.values.tolist()


gb_prem = gb_prem[gb_prem.iloc[:, 1].isin(gen_companies)].dropna(axis=1, how='all')
gb_prem.columns = header.values.tolist()
gb_prem.columns[3]

gb_prem.sort_values(by=gb_prem[0:5].columns[4], ascending=False, inplace=True)
data = gb_prem[gb_prem.columns[4]].head()
plt.hist(data.astype('int32'))
plt.show()
data.astype('int32')
# %%
"""
This function cleans up data and generates seperate data sets for insurance
and reinsurance data sets, regardless of whether it is Life or General
"""


def ReadData(sheet):
    # Specifies which worksheet to read
    df = pd.read_excel(excel_file, sheet_name=sheet)

    # Deletes rows that is fully NaN
    df.dropna(axis=0, how='all', inplace=True)
    # Deletes columns that is fully NaN
    df.dropna(axis=1, how='all', inplace=True)

    Data_Description = df.iloc[0, 0]  # Takes a specific data frame
    df = df.iloc[1:, :]  # Omits the first row from the data frame
    df.reset_index(drop=True, inplace=True)  # Reindexing the data set

    df.columns = df.iloc[0, :]  # Takes the first row of the data frame
    df = df.iloc[1:, :]
    # Retains any row not having the word Total
    df = df.loc[~df.Company.str.match('TOTAL')]

    # Any company name not having the name the, the first word is taken
    df.loc[~(df.Company.str.split().str[0] == 'THE'),
           'Adj_company'] = df.Company.str.split().str[0]
    # Any company name having the, the second name is taken
    df.loc[(df.Company.str.split().str[0] == 'THE'),
           'Adj_company'] = df.Company.str.split().str[1]

    df.index = df.iloc[:, -1]  # taking of the last column to be the index
    df = df.iloc[:, 1:]

    # Separating data up to row having the name Reinsurers
    df_ins = df.loc[:'REINSURERS'].copy()
    # Separating data from row having the name Reinsurers
    df_rei = df.loc['REINSURERS':].copy()

    # drop rows with 5 or more NaN
    df_ins.dropna(axis=0, thresh=5, inplace=True)
    # drop rows with 5 or more NaN
    df_rei.dropna(axis=0, thresh=5, inplace=True)

    df_ins.drop('Adj_company', axis=1, inplace=True)
    df_rei.drop('Adj_company', axis=1, inplace=True)

    df_x = [df_ins, df_rei]
    for i in df_x:
        for col in i.columns[0:(len(df.columns.values) - 1)]:
            i[col] = pd.to_numeric(i[col], errors='coerce')

    return df_ins
    return df_rei

# %%


GB_class_sheets = {'link': ['APPENDIX 13', 'APPENDIX 14', 'APPENDIX 15', 'APPENDIX 16', 'APPENDIX 17', 'APPENDIX 18'],
                   'desc': ['PREMIUM', 'MARKET_SHARE', 'CLAIMS_PAID', 'CLAIMS INCURRED', 'INCURRED_RATIO', 'UW_PROFIT']}
GB_class_sheets = pd.DataFrame(GB_class_sheets)

# This for Loop Merges all GB Classwise data into a single data frame

df_ins = ReadData(GB_class_sheets['link'][0])
GB_Classes = list(df_ins.columns[0:14].values)  # .values

df_ins.columns = pd.MultiIndex.from_product([[GB_class_sheets['desc'][0]],
                                             df_ins.columns], names=['PREMIUM', 'CLASS'])
for i in range(len(GB_class_sheets['link']) - 1):
    df = ReadData(GB_class_sheets['link'][i + 1])
    df.columns = pd.MultiIndex.from_product([[GB_class_sheets['desc'][i + 1]],
                                             df.columns], names=['PREMIUM', 'CLASS'])
    df_ins = pd.merge(df, df_ins, right_index=True, left_index=True)
del df

# %%
# Plotting of the graphs


df_ins.sort_values(by=[('PREMIUM', 'Motor Private')], ascending=False)[
    'PREMIUM']['Motor Private'][0:4].plot.bar()

# PLOT TOP 5 COMPANIES
# for i in GB_Classes:

df_ins.sort_values(by=[('PREMIUM', 'Total\r\n')],
                   ascending=False, inplace=True)

Spec_classes = ['Motor Private', 'Motor Commercial', 'Medical', 'Others']
S1 = df_ins.loc[:, ['PREMIUM', 'INCURRED_RATIO']]

a1 = list(S1['PREMIUM']['Motor Private'][0:5])
a2 = S1['PREMIUM']['Motor Commercial'][0:5]
a3 = S1['PREMIUM']['Medical'][0:5]
a4 = S1['PREMIUM'][GB_Classes].drop(Spec_classes[0:3], axis=1)[0:5].sum(axis=1)
a5 = S1['INCURRED_RATIO']['Total\r\n'][0:5]

S1.index[0:5]
S2 = pd.DataFrame(S2, columns=['Others'])
S = S1.join(S2)

fig, ax = plt.subplots()
plt.style.use('default')
y_pos = np.arange(5)
ax1 = ax.twinx()
p5 = ax1.plot(y_pos, a5, label='Overall LR')
p1 = ax.bar(y_pos, a1, align='center', label='Motor Private', alpha=0.85)
p2 = ax.bar(y_pos, a2, bottom=a1, label='Motor Commercial', alpha=0.85)
p3 = ax.bar(y_pos, a3, bottom=a1 + a2, label='Medical', alpha=0.85)
p4 = ax.bar(y_pos, a4, bottom=a1 + a2 + a3, label='Others', alpha=0.85)
plt.xticks(y_pos, S1.index[0:5])
fig.legend(loc="upper center")
plt.tight_layout()
plt.show()

Plot_data = df_ins.loc[:, [('UW_PROFIT', i),
                           ('PREMIUM', i),
                           ('INCURRED_RATIO', i)
                           ]][0:4]
G1 = ax.bar(Plot_data.index.values, Plot_data['PREMIUM']
            [i], label='PREMIUM', align='center', alpha=0.3)
G2 = ax.bar(Plot_data.index.values, Plot_data['UW_PROFIT']
            [i], label='UW_PROFIT', align='center', alpha=0.2)
G3 = ax1.plot(Plot_data.index.values,
              Plot_data['INCURRED_RATIO'][i], label='INCURRED_RATIO')
ax1.grid(None)
# leg=plt.legend()
fig.legend(loc="upper left")


for i in GB_Classes:
    fig, ax = plt.subplots()
    ax1 = ax.twinx()
    ax.set_title(i)
    df_ins.sort_values(by=[('PREMIUM', i)], ascending=False, inplace=True)
    Plot_data = df_ins.loc[:, [('UW_PROFIT', i),
                               ('PREMIUM', i),
                               ('INCURRED_RATIO', i)
                               ]][0:4]
    G1 = ax.bar(Plot_data.index.values, Plot_data['PREMIUM']
                [i], label='PREMIUM', align='center', alpha=0.3)
    G2 = ax.bar(Plot_data.index.values, Plot_data['UW_PROFIT']
                [i], label='UW_PROFIT', align='center', alpha=0.2)
    G3 = ax1.plot(Plot_data.index.values,
                  Plot_data['INCURRED_RATIO'][i], label='INCURRED_RATIO')
    ax1.grid(None)
    # leg=plt.legend()
    fig.legend(loc="upper left")

    plt.tight_layout()
    plt.show()


for i in range(2):
    plt.subplot(2, 1, i + 1)
    title = df_ins.columns[i]
    df_ins.sort_values(by=[title], ascending=False, inplace=True)
    Plot_data = df_ins[title][0:4]
    objects = Plot_data.index
    y_pos = np.arange(len(Plot_data.index))
    plt.bar(y_pos, Plot_data, align='center', alpha=0.5)
    plt.title('Top {}'.format(title))
    plt.xticks(y_pos, objects)

plt.tight_layout()
plt.show()
