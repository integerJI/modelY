from django.shortcuts import render
import io
from django.http import HttpResponse
from xlsxwriter.workbook import Workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook

def index(request):
    return render(request, 'index.html')

def download(request):
    sheetPath = request.GET['sheetPath']

    scope = [
        'https://spreadsheets.google.com/feeds'
        , 'https://www.googleapis.com/auth/drive'
    ]

    json_file_name = 'calm-magpie-382913-e094e6764e08.json'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
    gc = gspread.authorize(credentials)

    spreadsheet_url = sheetPath

    # 스프레스시트 문서 가져오기 
    doc = gc.open_by_url(spreadsheet_url)

    # 시트 선택하기
    worksheet = doc.worksheet('판매현황')

    # 열 데이터 가져오기
    column_data = worksheet.col_values(3)

    # 2차원 리스트로 변환
    column_data_2d = [[col] for col in column_data]

    # 엑셀 파일 생성
    output = io.BytesIO()
    workbook = Workbook()
    worksheet = workbook.active

    # 시트 데이터를 엑셀 파일에 저장
    for i, col in enumerate(column_data_2d):
        worksheet.cell(row=i+1, column=2, value=col[0])
    workbook.save(output)
    
    # response 생성
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=새로운 엑셀 파일명.xlsx'
    response.write(output.getvalue())

    return response
