import xlrd
import csv

input_excel_path = r'C:\Users\Bakhyt Kultay\Downloads\КАТО_17.11.2023.xls'
output_csv_path = 'КАТО_old.csv'


with xlrd.open_workbook(input_excel_path) as wb:
    ws = wb.sheet_by_index(0)

    with open(output_csv_path, 'w', newline='', encoding='utf') as csv_file:
        csv_writer = csv.writer(csv_file)

        for row in range(ws.nrows):
            csv_writer.writerow(ws.row_values(row))


with open(output_csv_path, 'r', newline='', encoding='utf') as csv_file:
    reader = csv.DictReader(csv_file)
    rows = list(reader)

    output_fields = reader.fieldnames
    output_fields.remove("nn")
    output_fields.remove("ab")
    output_fields.remove("cd")
    output_fields.remove("ef")
    output_fields.remove("hij")
    output_fields.remove("k")

    output_fields.extend(["id", "parent_id"])

    previous_i1, previous_i2 = None, None

    for i, row in enumerate(rows, start=1):
        row['id'] = i

        if float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) == 0 and float(row["hij"]) == 0:
            previous_i1 = i - 1
            row["parent_id"] = previous_i1
        elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) == 0:
            previous_i2 = i
            row["parent_id"] = previous_i1 + 1
        elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) > 0:
            row["parent_id"] = previous_i2
        else:
            row["parent_id"] = None

        if row['te'].endswith('.0'):
            row['te'] = str(int(float(row['te'])))


output_new_csv_path = 'КАТО_new.csv'
with open(output_new_csv_path, 'w', newline='', encoding='utf') as csv_new:
    writer = csv.DictWriter(
        csv_new,
        fieldnames=output_fields,
        extrasaction='ignore')

    writer.writeheader()
    for row in rows:
        writer.writerow(row)