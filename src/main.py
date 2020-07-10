# coding=gbk
"""
Author:      XuLin Yang
Student id:  904904
Date:        2020-7-10 22:38:17
Description: 
"""

import xlrd
import argparse
import pandas as pd
import numpy as np


def read_firm(firm_sheet: xlrd.sheet.Sheet):
    """
    Note row & column is 0-index based
    :param firm_sheet:
    :return:
    """
    # get firm name
    firm_name = firm_sheet.cell(7, 1).value

    # selected data column, 0 based index
    if firm_sheet.ncols == 10:
        # [存货编号, 存货全名, 采购数量, 价税合计]
        selected_columns = [1, 2, 3, 6]
    elif firm_sheet.ncols == 11:
        selected_columns = [1, 2, 3, 4, 7]
    else:
        print(firm_sheet.ncols, firm_sheet.name)
        raise Exception("Unsupported data sheet format")

    rows = []
    for i in range(20, firm_sheet.nrows):
        row = [firm_name]
        for j in selected_columns:
            entry = firm_sheet.cell(i, j).value

            try:
                entry = float(entry)
            except Exception:
                pass
            row.append(entry)

        has_empty_data = False
        # check empty entry in data
        for k in [-2, -1]:
            if row[k] == '':
                has_empty_data = True
                break
        if has_empty_data:
            continue

        # print(row)
        row.append(round(row[-1] / row[-2], 2))  # 含税单价 = 价税合计 / 采购数量
        rows.append(row)

    data_frame = pd.DataFrame(rows, columns=["公司名称"] + [firm_sheet.cell(19, i).value for i in selected_columns] + ["含税单价"])
    print(data_frame)
    if "基本单位" not in data_frame.columns:
        data_frame.insert(3, "基本单位", np.nan, True)
    return data_frame


def read_excel(file_path: str):
    wb = xlrd.open_workbook(filename=file_path)
    n_firm = len(wb.sheets())

    res = []
    for i in range(0, n_firm):
        res.append(read_firm(wb.sheet_by_index(i)))
    return res


def difference(row):
    return round(row["含税单价_x"] - row["含税单价_y"], 2)


def write_excel(firms_data: list):
    # wb = xlwt.Workbook(encoding="utf-8")
    n_firm = len(firms_data)
    # write_to = wb.add_sheet('Sheet {}'.format(n_firm+1), cell_overwrite_ok=True)
    #
    # # find column names in firms data
    # cur = firms_data[0].columns
    # for firm in firms_data:
    #     if len(firm.columns) > len(cur):
    #         cur = firm.columns
    #
    # # write head row
    # print(cur)
    # n = len(cur)
    # for i, c in enumerate(cur):
    #     write_to.write(0, i, c)
    #     write_to.write(0, i + n, c)

    # join firm
    res = None
    has_result = False

    for i in range(0, n_firm):
        firm1 = firms_data[i]
        for j in range(i+1, n_firm):
            firm2 = firms_data[j]

            merged = pd.merge(firm1, firm2, on=["存货全名"], how='inner')
            if not res:
                res = merged
                has_result = True
            else:
                pd.concat(res, merged)

    if not has_result:
        res = pd.DataFrame(columns=["啥都没有"])
    else:
        res["含税单价差价 (x-y)"] = res.apply(lambda row: difference(row), axis=1)
        res["含税单价差价 * 公司X)"] = res.apply(lambda row: int(row["含税单价差价 (x-y)"] * row["采购数量_x"]), axis=1)
        res["含税单价差价 * 公司Y)"] = res.apply(lambda row: int(row["含税单价差价 (x-y)"] * row["采购数量_y"]), axis=1)
    res.to_excel("结果.xls")
    # wb.save("./result.xls")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='firm products comparison form excel')
    parser.add_argument('-f', help='excel file path')

    args = parser.parse_args()

    data = read_excel(args.f)
    print(data)
    write_excel(data)
