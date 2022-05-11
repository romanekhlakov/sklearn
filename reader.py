import pandas as pd
import os

from datetime import datetime


def excel_parse(path='data'):
    xls = pd.ExcelFile(path)

    sheet_to_df_map = {}
    for sheet_name in xls.sheet_names:
        date = datetime.strptime(sheet_name, '%m-%Y')
        sheet_to_df_map[date] = xls.parse(sheet_name).iloc[5, 5]

    df = pd.DataFrame(sheet_to_df_map.items()).rename(columns={0: 'DateTime', 1: 'Signal'})
    df = df.set_index('DateTime')

    return df


def parse(path='data'):
    dic = {}
    for i in os.listdir(path):
        try:
            cwd = os.getcwd()
            path_dir = os.path.join(cwd, path + '/' + i)
            dic[i] = excel_parse(path_dir)
        except:  # ignore .gitignore
            pass

    list_i = []
    for i in dic:
        list_i.append(dic[i])

    return pd.concat(list_i).sort_index()
