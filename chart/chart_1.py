import xlsxwriter

workbook = xlsxwriter.Workbook('chart.xlsx')
worksheet = workbook.add_worksheet()


chart = workbook.add_chart({'type': 'radar'})

data = [
    [1, 2, 3, 4, 5],
    [2, 4, 6, 8, 10],
    [3, 6, 9, 12, 15],
]

worksheet.write_column('A1', data[0])
worksheet.write_column('B1', data[1])
worksheet.write_column('C1', data[2])

chart.add_series({'values': '=Sheet1!$A$1:$A$5'})


worksheet.insert_chart('A7', chart)

workbook.close()