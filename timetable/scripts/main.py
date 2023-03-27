import openpyxl
import json
import string
import time
from collections import OrderedDict


def alphabet_to_number(column):
    alphabets = string.ascii_uppercase
    # colmn はインデックスの何番目かを返す
    return alphabets.index(column) + 1


def output_json(import_path):
    # Excelファイルを読み込みます
    workbook = openpyxl.load_workbook(import_path, read_only=True, data_only=True)
    worksheet = workbook['全科目']

    min_col = alphabet_to_number("A")
    max_col = alphabet_to_number("M")
    min_row = 14

    timetables = []
    for row in worksheet.iter_rows(values_only=True, min_row=min_row, min_col=min_col, max_col=max_col):
        timetable_dict = {
            "No": str(row[0]),
            "gakubu": row[1],
            "gakka": row[2],
            "niti": row[3],
            "gen": row[4],
            "gakki": row[5],
            "kurasu": row[6],
            "kamoku": row[7],
            "tantou": row[8].replace('\u3000', ' '),
            "kyoushitsu": row[9],
            "sou": row[10],
            "ji": row[11],
            "bikou": str(row[12])
        }

        timetables.append(timetable_dict)

        # print(timetable_dict)
    # JSONファイルに書き込みます
    with open('output.json', 'w', encoding='utf-8') as outfile:
        json.dump(timetables, outfile, ensure_ascii=False)

def main():
    start_time = time.time()
    output_json('2023_jikanwari_kensaku.xlsx')
    end_time = time.time()
    elapsed_time = end_time - start_time
    print("Elapsed time: {:.2f} seconds".format(elapsed_time))

if __name__ == "__main__":
    main()