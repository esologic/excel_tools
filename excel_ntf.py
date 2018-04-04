from openpyxl import load_workbook


class ExcelToNTF(object):

    def __init__(self, excel_path, ntf_path):

        with open(ntf_path, 'w') as ntf_file:

            ntf_file.write("# NTF file generated automatically\n")

            workbook = load_workbook(excel_path)

            for sheet_number, sheet in enumerate(workbook.worksheets):

                for row_number, row in enumerate(sheet.rows):

                    for cell_number, cell in enumerate(row):

                        node_id = cell.value

                        if node_id is not None:

                            x = cell_number
                            y = row_number
                            z = sheet_number

                            line_text = str(node_id) + "," + str(x) + "," + str(y) + "," + str(z) + "\n"
                            ntf_file.write(line_text)
