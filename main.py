import os
import sys
import xlrd # version=1.2.0

# create direction if not exists
def create_direction(output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

# convert xlsx to markdown table
def convert_csv_to_markdown_table(csv_file):
    # read xlsx file
    wb = xlrd.open_workbook(csv_file)
    for sheet_name in wb.sheet_names():
        sheet = wb.sheet_by_name(sheet_name)
        markdown_table = ''
        # read excel header
        markdown_table += '| ' + ' | '.join([str(cell.value) for cell in sheet.row(0)]) + ' |\n'
        markdown_table += '|:---' * sheet.ncols + '|\n'
        # read excel body
        for row in range(1, sheet.nrows):
            markdown_table += '| ' + ' | '.join([str(cell.value) for cell in sheet.row(row)]) + ' |\n'
        # write to md file
        with open(f'{output_dir}/{sheet_name}.md', 'w') as f:
            f.write(markdown_table)

# main
def main():
    create_direction(output_dir)
    convert_csv_to_markdown_table(csv_file)

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python mixed.py <PATH_TO_CSV_FILE> <OUTPUT_DIR>')
        sys.exit(1)
    csv_file = sys.argv[1]
    output_dir = sys.argv[2]
    main()
