#!/usr/bin/env python3

'''
Author : Md.Nazmul Hasan
Date   : 2020-05-03
Description : This script used for generate a result in xlsx file from skype-message JSON data
usage:
./json-to-xlsx.py
python json-to-xlsx.py
'''

import os
import json
import time
import pandas as pd

start_time = int(round(time.time() * 1000))


def main():
    out = os.path.abspath('./output')
    data = os.path.abspath('./data')
    json_file = 'skype-messages.json'
    input_file = open(data+os.path.sep+json_file)
    filename = 'Skype_Message.xlsx'
    json_array = json.load(input_file)
    sheets = []
    for item in json_array:
        sheets.append(item['group'])
    writer = pd.ExcelWriter(out+os.path.sep+filename)
    for sheet in sheets:
        print('Writing on sheet', sheet)
        desired_data = april_data(sheet, json_array)
        df = pd.DataFrame(desired_data)
        output = df.groupby(["date", "user"])["content"].count().reset_index().pivot("user", "date").fillna("")
        output.to_excel(writer, sheet)
    print('Generate successful !')
    writer.save()


def april_data(sheet, json_array):
    date_list = []
    for obj in json_array:
        if sheet == obj['group']:
            objects = obj['messages']
            for out in objects:
                desired_date = out['originalarrivaltime']
                full_date = desired_date[0:10]
                out['date'] = full_date
                out['user'] = out['displayName']
                ym = full_date[0:7]
                values = ym.split('-')
                if values[0] == '2020' \
                        and values[1] == '04' \
                        and len(out['content']) > 0 \
                        and out['displayName'] is not None:
                    date_list.append(out)
    return date_list


def finish():
    end_time = int(round(time.time() * 1000))
    print("Script took total time {} seconds".format((end_time - start_time) / 1000))


def run():
    main()
    finish()


if __name__ == '__main__':
    run()
