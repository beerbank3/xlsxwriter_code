import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Format01.xlsx')
worksheet = workbook.add_worksheet()

cell_format = workbook.add_format({
    'bold': True,
    'font_size': 12,
})
cell_color_format = workbook.add_format({
    'color': 'red',
})
cell_center_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
})
cell_bg_format = workbook.add_format({
    'bg_color': 'yellow',
})
worksheet.write       (0, 0, 'Foo', cell_format)
worksheet.write_string(1, 0, 'Bar', cell_center_format)
worksheet.write_number(2, 0, 3,     cell_color_format)
worksheet.write_blank (3, 0, '',    cell_bg_format)

workbook.close()