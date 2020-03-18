import pandas as pd
from functools import reduce


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


def sheet_names(year, quarter=4):
    """extract list of available sheets and their categories"""
    excel_file = '/Users/williesmboko/Desktop/Python/IRA analysis/' + \
        str(year) + '_Q'+str(quarter)+'_Statistics.xlsx'  # link to your file
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


def read_data(sheet, year=2018, quarter=4):
    """reads data given a specified sheet"""
    excel_file = '/Users/williesmboko/Desktop/Python/IRA analysis/' + \
        str(year) + '_Q'+str(quarter)+'_Statistics.xlsx'  # link to your file
    df = pd.read_excel(excel_file, sheet_name=sheet)
    return df


def get_companies(sheet_number, x_sheets, year=2018, quarter=4):
    """ extract list of companies"""
    rem_list = ['TOTAL', 'REINSURERS',
                'GRAND TOTAL', 'Amounts in Thousand Shillings']
    df = read_data(x_sheets['Link'][sheet_number], year, quarter)
    ins_names = list(df.iloc[4:, 1])
    reins_names = list(filter(lambda x: 'REINS' in x, ins_names))
    rem_list.extend(reins_names)
    ins_names = list(filter(lambda x: x not in rem_list, ins_names))
    rem_list = ['TOTAL', 'REINSURERS',
                'GRAND TOTAL', 'Amounts in Thousand Shillings']
    reins_names = list(filter(lambda x: x not in rem_list, reins_names))
    return ins_names, reins_names


def extract_pl_table(year, quarter, sheet, x_companies):
    """Extracts profit and loss tables from the sheets,\
    x_company --> (gen_companies,life_companies)"""
    df = read_data(sheet, year, quarter)
    header = df[df.iloc[:, 1] == 'Company'].dropna(axis=1)
    df = df[df.iloc[:, 1].isin(
        x_companies)].dropna(axis=1, how='all')
    df.columns = header.values[0].tolist()
    df.columns = [x.strip() for x in df.columns]
    df['Company'] = df['Company'].apply(strip_name)
    df['Company'] = df['Company'].apply(lambda x: x.strip())
    df.set_index('Company', inplace=True)
    for col in range(len(df.columns[1:])):  # converts all numbers to type float
        #df.iloc[:, col] = df.iloc[:, col].astype('float')
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col], errors='coerce')
    return df


def extract_bs_table(year, quarter, sheet, x_companies):
    """Extracts profit and loss tables from the sheets,\
    x_company --> (gen_companies,life_companies)"""
    df = read_data(sheet, year, quarter)
    header = df[df.iloc[:, 1] == 'Company'].dropna(axis=1)
    df = df.dropna(axis=1, how='all')
    header = header.values[0]
    df = df.iloc[:, range(len(header))]
    df.columns = [strip_name(x).strip() for x in header]
    df.dropna(how='any', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.drop(index=[0], inplace=True)
    df = df.T
    df.columns = df.iloc[0, :].values
    df.drop(axis=0, index='Company', inplace=True)
    for col in df.columns[1:]:  # converts all numbers to type float
        df[col] = df[col].astype('float', errors='ignore')
    return df


# # TODO : compare sheets in different years and highlight new companies
gen_sheets_2019, life_sheets_2019, bal_sheets_2019, pnl_sheets_2019 = \
    sheet_names(2019, 3)
gen_companies, gen_companies_rein = get_companies(0, gen_sheets_2019, 2019, 3)
life_companies, life_companies_rein = get_companies(1, life_sheets_2019, 2019, 3)

ira_data = {}
for y in range(2018, 2020):
    if y not in ira_data:
        ira_data[y] = {}
        for q in range(1, 5):
            if q not in ira_data[y]:
                ira_data[y][q] = {}
                ira_data[y][q]["general"] = {}
                ira_data[y][q]["life"] = {}

                ira_data[y][q]["general"]["class_data"] = {}
                ira_data[y][q]["general"]["profit_loss_account"] = {}
                ira_data[y][q]["general"]["bal_sheet_account"] = {}

                ira_data[y][q]["general"]["class_data"]["premium"] = extract_pl_table(
                    y, q, "APPENDIX 13", gen_companies)
                ira_data[y][q]["general"]["class_data"]["market_share"] = extract_pl_table(
                    y, q, "APPENDIX 14", gen_companies)
                ira_data[y][q]["general"]["class_data"]["loss_ratio"] = extract_pl_table(
                    y, q, "APPENDIX 17", gen_companies)
                ira_data[y][q]["general"]["class_data"]["claims_paid"] = extract_pl_table(
                    y, q, "APPENDIX 15", gen_companies)
                ira_data[y][q]["general"]["class_data"]["underwriting_profits"] = extract_pl_table(
                    y, q, "APPENDIX 18", gen_companies)
                ira_data[y][q]["general"]["class_data"]["claims_incurred"] = extract_pl_table(
                    y, q, "APPENDIX 16", gen_companies)
                # import P&L account
                ira_data[y][q]["general"]["profit_loss_account"]['p&l'] = extract_pl_table(
                    y, q, 3, gen_companies)  # problem importing 'APPENDIX 1'
                ira_data[y][q]["general"]["profit_loss_account"]['revenue'] = extract_pl_table(
                    y, q, "APPENDIX 19", gen_companies)
                # import balance sheet account
                ira_data[y][q]["general"]["bal_sheet_account"] =\
                    reduce(lambda left, right: pd.concat([left, right], sort=False),
                           [extract_bs_table(y, q, i, gen_companies) for i in [-4, -3, -2, -1]])

                ira_data[y][q]["life"]["class_data"] = {}
                ira_data[y][q]["life"]["profit_loss_account"] = {}
                ira_data[y][q]["life"]["bal_sheet_account"] = {}
