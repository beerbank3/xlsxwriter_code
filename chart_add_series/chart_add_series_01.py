import xlsxwriter

# 1. 워크북 및 워크시트 생성
workbook = xlsxwriter.Workbook('chart_add_series_example.xlsx')
worksheet = workbook.add_worksheet()

# 2. 차트 데이터 생성
worksheet.write('A1', '날짜')
worksheet.write('B1', '데이터1')
worksheet.write('C1', '데이터2')

data = [
    ['2023-01-01', 10, 20],
    ['2023-01-02', 15, 25],
    ['2023-01-03', 20, 30],
    ['2023-01-04', 25, 35],
    ['2023-01-05', 30, 40],
    ['2023-01-06', 35, 45],
    ['2023-01-07', 40, 50],
]

for row_num, row_data in enumerate(data):
    worksheet.write_row(row_num + 1, 0, row_data)

# 3. 차트 객체 생성
line_chart = workbook.add_chart({'type': 'line'})

# 4. 데이터 시리즈 추가
line_chart.add_series({
    'name':       'chart_1',
    'categories': '=Sheet1!$A$2:$A$8',  # X 축 범주
    'values':     '=Sheet1!$B$2:$B$8',  # Y 축 값
    'marker':     {
        'type': 'diamond', # 마커 모양
        'size': 10,        # 마커 크기
        'border': {'color': 'black'}, # 마커 테두리 색상
        'fill':   {'color': 'red'},   # 마커 채우기 색상
        },
})

# 5. 차트에 추가적인 시리즈 추가 가능
column_chart = workbook.add_chart({'type': 'column'})
column_chart.add_series({
    'name':       'chart_2',
    'categories': '=Sheet1!$A$2:$A$8',
    'values':     '=Sheet1!$C$2:$C$8',
    'bg_color':   'yellow',
})

column_chart.combine(line_chart) # line_chart를 column_chart에 추가
column_chart.set_title({
    'name': '차트', # 차트 제목
    'name_font': {'size': 14, 'bold': True}, # 차트 제목 폰트
    'overlay': True, # 차트 제목이 차트 영역을 덮는지 여부
    'layout': {
        'x': 0.5, # 차트 제목 x 좌표
        'y': 0.5, # 차트 제목 y 좌표
    },
    }) 
column_chart.set_legend({'position': 'bottom'}) # 기본값은 right (차트 범례 위치)
column_chart.set_x_axis({'name': '날짜'}) # X 축 이름
column_chart.set_y_axis({'name': '데이터'}) # Y 축 이름

column_chart.set_style(37) # 차트 스타일 설정
column_chart.set_size({'width': 720, 'height': 576}) # 차트 크기 설정
column_chart.set_chartarea({'border': {'color': 'red'}}) # 차트 영역 테두리 색상 설정
column_chart.set_plotarea({
    'gradient': {'colors': ['#FFEFD1', '#F0EBD5', '#D8D0C9']}, # 차트 플롯 영역 그라데이션 색상 설정
    'border':   {'color': 'red'}, # 차트 플롯 영역 테두리 색상 설정
    'fill':     {'color': 'green'}, # 차트 플롯 영역 채우기 색상 설정
    })

column_chart.set_table() # 차트 테이블 설정
column_chart.set_up_down_bars() # 차트 상하 막대 설정
column_chart.set_drop_lines() # 차트 드롭 라인 설정
column_chart.set_high_low_lines() # 차트 최고 최저선 설정
# 6. 차트를 워크시트에 삽입
worksheet.insert_chart('E2', column_chart)

# 7. 워크북 닫기
workbook.close()