import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
=======
from functools import reduce
plt.rcdefaults()

def sheet_names(year):
    """extract list of available sheets and their categories"""
    excel_file = '/Users/williesmboko/Desktop/Python/IRA analysis/' + \
        str(year) + '_Q3_Statistics.xlsx'  # link to your file
    sheet_names = pd.read_excel(excel_file, sheet_name='Table of Contents',
                                header=6, usecols=[1, 2])
    sheet_names['Link'] = sheet_names['Link'].apply(lambda x: x.strip("''"))
    # identifies different report types
    gen_sheets = sheet_names[sheet_names['Description'].str.contains(
        'GENERAL')]
    life_sheets = sheet_names[~sheet_names['Description'].str.contains(
        'GENERAL')]
    bal_sheets = sheet_names[sheet_names['Description'].str.contains(
        'BALANCE')]
    pnl_sheets = sheet_names[~sheet_names['Description'].str.contains(
        'BALANCE')]
    return gen_sheets, life_sheets, bal_sheets, pnl_sheets


def read_data(year, sheet):
    """reads data given a specified sheet"""
    excel_file = '/Users/williesmboko/Desktop/Python/IRA analysis/' + \
        str(year) + '_Q3_Statistics.xlsx'  # link to your file
    df = pd.read_excel(excel_file, sheet_name=sheet)
    return df

def get_companies(year, sheet_number, x_sheets):
    """ extract list of companies"""
    rem_list = ['TOTAL', 'REINSURERS',
                'GRAND TOTAL', 'Amounts in Thousand Shillings']
    df = read_data(year, x_sheets['Link'][sheet_number])
    ins_names = list(df.iloc[4:, 1])
    reins_names = list(filter(lambda x: 'REINS' in x, ins_names))
    rem_list.extend(reins_names)
    ins_names = list(filter(lambda x: x not in rem_list, ins_names))
    rem_list = ['TOTAL', 'REINSURERS',
                'GRAND TOTAL', 'Amounts in Thousand Shillings']
    reins_names = list(filter(lambda x: x not in rem_list, reins_names))
    return ins_names, reins_names


def extract_pl_table(year, sheet, x_companies):
    """Extracts profit and loss tables from the sheets,\
    x_company --> (gen_companies,life_companies)"""
    df = read_data(year, sheet)
    header = df[df.iloc[:, 1] == 'Company'].dropna(axis=1)
    df = df[df.iloc[:, 1].isin(
        x_companies)].dropna(axis=1, how='all')
    df.columns = header.values[0].tolist()
    df.reset_index(drop=True, inplace=True)
    for col in df.columns[1:]:  # converts all numbers to type float
        df[col] = df[col].astype('float')
    return df


def strip_name(name):
    """Simplifies company display name """
    denied_names = ['INSURANCE', 'INSURANE', 'COMPANY', 'GENERAL', 'THE',
                    'ASSURANCE', 'LION', 'SHIELD', 'KENYA']
    sep_names = name.split(" ")
    sep_names_fltd = list(
        filter(lambda name: name not in denied_names, sep_names))
    final_word = ""
    for i in range(len(sep_names_fltd)):
        final_word += sep_names_fltd[i]
        final_word += " "
    return final_word


def millions(x, pos):
    """The two args are the value and tick position"""
    return '%1.2fB' % (x * 1e-6)


def auto_label(rects):
    """Attach a text label above each bar in *rects*, displaying its height."""
    for rect in rects:
        height = rect.get_height()
        height_display = '%1.2fB' % (height * 1e-6)
        ax.annotate('{}'.format(height_display),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 1),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')


# TODO : compare sheets in different years and highlight new companies
gen_sheets_2019, life_sheets_2019, bal_sheets_2019, pnl_sheets_2019 = \
    sheet_names(2019)
gen_companies, gen_companies_rein = get_companies(2019, 0, gen_sheets_2019)
life_companies, life_companies_rein = get_companies(2019, 1, life_sheets_2019)


# extract data here with these extract_pl_table
prem_2018 = extract_pl_table(2018, "APPENDIX 13", gen_companies)
prem_2019 = extract_pl_table(2019, "APPENDIX 13", gen_companies)
market_share_2018 = extract_pl_table(2018, "APPENDIX 14", gen_companies)
market_share_2019 = extract_pl_table(2019, "APPENDIX 14", gen_companies)
loss_ratio_2018 = extract_pl_table(2018, "APPENDIX 17", gen_companies)
loss_ratio_2019 = extract_pl_table(2019, "APPENDIX 17", gen_companies)

# simplify company names
for sheets in [prem_2018, prem_2019,
               market_share_2018, market_share_2019,
               loss_ratio_2018, loss_ratio_2019]:
    sheets['Company'] = sheets['Company'].apply(strip_name)


names = ['Company',
         '2018_premium',
         '2019_premium',
         '2018_market share',
         '2019_market share',
         '2018_loss ratio',
         '2019_loss ratio']
prem_2019.columns[i]
# TODO insert for loop here
for i in range(8, 9):
    plot_data = reduce(lambda left, right: left.merge(right, on='Company'),
                       [
        prem_2018[['Company', prem_2018.columns[i]]],
        prem_2019[['Company', prem_2018.columns[i]]],
        market_share_2018[['Company', prem_2018.columns[i]]],
        market_share_2019[['Company', prem_2018.columns[i]]],
        loss_ratio_2018[['Company', prem_2018.columns[i]]],
        loss_ratio_2019[['Company', prem_2018.columns[i]]],
    ])
    plot_data.columns = names

    # plt style
    plot_data.sort_values(by='2019_premium', ascending=False, inplace=True)
    plot_data = plot_data[plot_data['2018_premium'] > 0]
    plot_data.reset_index(drop=True, inplace=True)
    color1 = '#fec615'
    color2 = '#800020'

    # oop style
    fig, ax = plt.subplots(2, 2, figsize=(12, 12))
    x_pos = np.arange(1, 6)
    ax[0, 0].set_title('gross premium', fontweight='bold')
    width = 0.4
    ax[0, 0].bar(x_pos - width/2,
                 np.array(plot_data['2018_premium'].head(5))/1e6,
                 width=0.4, alpha=1, label='2018')
    ax[0, 0].bar(x_pos + width/2,
                 np.array(plot_data['2019_premium'].head(5))/1e6,
                 width=0.4, alpha=1, label='2019')
    for index, value in enumerate(plot_data['2018_premium'].head(5)):
        ax[0, 0].text(1 + index - width/2, 0,
                      str(round(value / 1e6, 2)),
                      ha='center', va='bottom', fontsize=9,
                      fontweight='bold', color='w')
    for index, value in enumerate(plot_data['2019_premium'].head(5)):
        ax[0, 0].text(1 + index + width/2, 0,
                      str(round(value / 1e6, 2)), ha='center',
                      va='bottom', fontsize=9, fontweight='bold', color='w')
    for index, value in enumerate(plot_data['2018_premium'].head(5)):
        ax[0, 0].text(1 + index - width/2, value/1e6,
                      str(round(plot_data['2018_market share'][index], 1))+'%',
                      ha='center', va='bottom',
                      fontsize=7.5, fontweight='bold')
    for index, value in enumerate(plot_data['2019_premium'].head(5)):
        ax[0, 0].text(1 + index + width/2, value/1e6,
                      str(round(plot_data['2019_market share'][index], 1))+'%',
                      ha='center', va='bottom',
                      fontsize=7.5, fontweight='bold')
    ax[0, 0].set_xticks(x_pos)
    ax[0, 0].set_xticklabels(plot_data['Company'].head(5))
    ax[0, 0].set_ylabel('KES B')
    ax[0, 0].legend()

    ax[1, 0].plot(plot_data['2018_loss ratio'].rank(pct=True, ascending=False),
                  plot_data['2019_loss ratio'].rank(pct=True, ascending=False),
                  color='b', marker='o', linestyle='None')
    ax[1, 0].set_title('loss ratio rank', fontweight='bold')
    ax[1, 0].set_xlabel('bottom <-- 2018 rank  --> top')
    ax[1, 0].set_ylabel('bottom <-- 2019 rank  --> top')
    ax[1, 0].vlines([0.5], 0, 1, transform=ax[1, 0].get_xaxis_transform(), color='r')
    ax[1, 0].hlines([0.5], 0, 1, transform=ax[1, 0].get_yaxis_transform(), color='r')
    for index, value in enumerate(plot_data['2019_loss ratio'].rank(pct=True, ascending=False).head(5)):
        ax[1, 0].text(plot_data['2018_loss ratio'].rank(pct=True, ascending=False)[index], value,
                      plot_data['Company'][index], ha='center', va='bottom', color='g')
    plt.setp(ax[1, 0].get_xticklabels(), visible=False)
    plt.setp(ax[1, 0].get_yticklabels(), visible=False)

    ax[0, 1].set_title('loss ratio', fontweight='bold')
    ax[0, 1].plot(plot_data['2018_loss ratio'].head(5),
                  marker='o', linestyle='--', label='2018')
    ax[0, 1].plot(plot_data['2019_loss ratio'].head(5),
                  marker='x', linestyle='--', label='2019')
    ax[0, 1].legend()
    plt.tight_layout()
    plt.show()
