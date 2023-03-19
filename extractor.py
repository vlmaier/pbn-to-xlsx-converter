import argparse
import xlsxwriter


def extract(file_name):
    data = [line for line in open(file_name, 'r') if '[DoubleDummyTricks' in line]
    xlsx_file_name = file_name.replace('.pbn', '.xlsx')
    workbook = xlsxwriter.Workbook(xlsx_file_name)
    worksheet = workbook.add_worksheet()
    row = 0
    for line in data:
        line = line\
            .replace('\n', '')\
            .replace('[', '')\
            .replace(']', '')\
            .replace('"', '')\
            .replace('DoubleDummyTricks', '')\
            .strip()
        line = ','.join(line)\
            .replace(',', '', 0)\
            .replace('a', '10')\
            .replace('b', '11')\
            .replace('c', '12')\
            .replace('d', '13')
        col = 0
        for entry in line.split(','):
            worksheet.write(row, col, int(entry))
            col += 1
        row += 1
    workbook.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", "-f", type=str, required=True)
    args = parser.parse_args()
    extract(args.file)
