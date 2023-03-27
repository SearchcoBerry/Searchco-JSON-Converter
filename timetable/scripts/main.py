import openpyxl
import json
import string
import time


def alphabet_to_number(column):
    alphabets = string.ascii_uppercase
    # colmn はインデックスの何番目かを返す
    return alphabets.index(column) + 1


def output_json(import_path):
    # Excelファイルを読み込みます
    workbook = openpyxl.load_workbook(import_path, read_only=True, data_only=True)
    worksheet = workbook['全科目']

    data = []

    min_col = alphabet_to_number("A")
    max_col = alphabet_to_number("M")
    min_row = 13

    for row in worksheet.iter_rows(values_only=True, min_row=min_row, min_col=min_col, max_col=max_col):
        row_data = []
        print(row)
        """
        for cell in row:
            row_data.append(cell.value)
            print(cell.value)
        # data.append(row_data)
        """

    print(f'最終行: {worksheet.max_row}')

def main():
    start_time = time.time()
    output_json('2023_jikanwari_kensaku.xlsx')
    end_time = time.time()
    elapsed_time = end_time - start_time

if __name__ == "__main__":
    main()