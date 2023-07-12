from datetime import timedelta
from datetime import datetime

from flask import Flask, request, render_template, send_file
import pandas as pd
from openpyxl import Workbook
from io import BytesIO
import os
import re
from datetime import time

app = Flask(__name__)

# 결과 파일을 저장할 폴더 경로
RESULT_FOLDER = 'result'


@app.route('/overwork', methods=['GET', 'POST'])
def upload_over_work_file():
    if request.method == 'POST':
        # 업로드된 파일 가져오기
        work_file = request.files['file1']
        over_work_file = request.files['file2']

        # xlsx 파일 읽기
        df = pd.read_excel(work_file, sheet_name="result_row")
        df_over_work = pd.read_excel(over_work_file)

        # xlsx 파일 처리
        df_original, df_result = process_overwork_xlsx(df, df_over_work)

        # 결과 파일을 로컬에 저장
        if not os.path.exists(RESULT_FOLDER):
            os.makedirs(RESULT_FOLDER)

        # 다운로드할 파일 생성
        original_filename = over_work_file.filename  # 업로드된 파일의 원래 파일명
        filename, extension = os.path.splitext(original_filename)  # 파일명과 확장자 분리
        result_filename = f"{filename}_result{extension}"  # 다운로드할 파일의 이름 생성

        result_filepath = os.path.join(RESULT_FOLDER, result_filename)

        # XLSX 파일 생성
        with pd.ExcelWriter(result_filepath, engine='xlsxwriter') as writer:

            # 각 데이터프레임을 시트로 저장
            df_original.to_excel(writer, sheet_name='original', index=False)
            df_result.to_excel(writer, sheet_name='result', index=False)

        df_html = df_result.to_html()

        return render_template('overwork.html', df_html=df_html, over_work_file_download=result_filepath)

    return '''
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file1">근무기록 Result 파일 업로드<br><br>
      <input type="file" name="file2">야근신청서<br><br>
      <input type="submit" value="업로드">
    </form>
    '''


def process_overwork_xlsx(df, df_overwork):
    df_overwork['결과'] = ''
    df_overwork['날짜'] = pd.to_datetime(df_overwork['야간 근무 일자']).dt.date

    df_overwork['결과'] = '확인필요'

    df_overwork['총근무시간'] = ''

    # df_overwork 데이터프레임을 순회하면서 df와 조건 일치 여부 확인
    for index, row in df_overwork.iterrows():
        name = row['이름']
        date = row['날짜']
        mask = (df['이름'] == name) & (pd.to_datetime(df['날짜']).dt.date == date)
        if len(df[mask]) > 0:
            work_time = df.loc[mask, '총근무시간'].values[0]
            hour_str, minute_str = work_time.split('시간 ')

            hour = int(hour_str)
            minute = int(re.sub(r'\D', '', minute_str))

            total_minutes = hour * 60 + minute

            if total_minutes >= 10 * 60:
                df_overwork.loc[index, '결과'] = '정상'

            if row['상태'] == '취소':
                df_overwork.loc[index, '결과'] = '상신취소'

            df_overwork.loc[index, '총근무시간'] = work_time

    df_overwork.drop('날짜', axis=1, inplace=True)

    df_result = pd.DataFrame()
    df_result['문서 번호'] = df_overwork['문서 번호']
    df_result['이름'] = df_overwork['이름']
    df_result['야간 근무 일자'] = df_overwork['야간 근무 일자']
    df_result['총근무시간'] = df_overwork['총근무시간']
    df_result['결과'] = df_overwork['결과']

    return df_overwork, df_result


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # 업로드된 파일 가져오기
        uploaded_file = request.files['file']

        # xlsx 파일 읽기
        df = pd.read_excel(uploaded_file)

        # xlsx 파일 처리
        df_result, df_result_row = process_xlsx(df)

        # 결과 파일을 로컬에 저장
        if not os.path.exists(RESULT_FOLDER):
            os.makedirs(RESULT_FOLDER)

        # 다운로드할 파일 생성
        original_filename = uploaded_file.filename  # 업로드된 파일의 원래 파일명
        filename, extension = os.path.splitext(original_filename)  # 파일명과 확장자 분리
        result_filename = f"{filename}_result{extension}"  # 다운로드할 파일의 이름 생성

        result_filepath = os.path.join(RESULT_FOLDER, result_filename)

        # XLSX 파일 생성
        writer = pd.ExcelWriter(result_filepath, engine='xlsxwriter')

        # 각 데이터프레임을 시트로 저장
        df.to_excel(writer, sheet_name='Original', index=False)
        df_result.to_excel(writer, sheet_name='result', index=False)
        df_result_row.to_excel(writer, sheet_name='result_row', index=False)

        writer.close()

        df_html = df_result.to_html()

        return render_template('index.html', df_html=df_html, file_download=result_filepath)

    return '''
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="upload">
    </form>
    '''


@app.route('/download_overwork')
def download_overwork():
    output = request.args.get('over_work_file_download', None)
    return send_file(output, as_attachment=True)


@app.route('/download_excel')
def download_excel():
    output = request.args.get('file_download', None)
    return send_file(output, as_attachment=True)


# 근무 시간 계산 함수
def calculate_working_hours(row):
    # 시작시각이 종료시각 보다 클 경우, 철야 근무 한 것으로 판단하여 계산 & 철야는 휴게시간 2 시간 적용
    if row['종료시각'] < row['시작시각']:
        working_hours = row['종료시각'] + timedelta(days=1) - row['시작시각'] - timedelta(hours=2)

    #  timedelta(hours=1) 한 시간 씩 빼는 것은 휴게시간을 1시간으로 가정하여 일괄 마이너스 처리
    else:
        working_hours = row['종료시각'] - row['시작시각'] - timedelta(hours=1)
    return working_hours


def process_xlsx(df):
    # 1. 결측값 제거
    # 조직 or 근무 정책 컬럼에 '수원'이 포함되었을 경우
    df = df.dropna(subset=['조직'])
    df = df[~df['근무유형'].str.contains('수원')]

    # 주말에 해당하는 날짜의 데이터는 제거
    # 날짜 컬럼을 datetime 형식으로 변환
    df['날짜'] = pd.to_datetime(df['날짜'])

    # 주말에 해당하는 토요일(Saturday) 및 일요일(Sunday)을 제거
    df = df[~df['날짜'].dt.dayofweek.isin([5, 6])]

    # 시작 시간을 datetime 형식으로 변환
    df['시작시각'] = pd.to_datetime(df['시작시각'], format='%H:%M', errors='coerce')

    # 종료 시간을 datetime 형식으로 변환
    df['종료시각'] = pd.to_datetime(df['종료시각'], format='%H:%M', errors='coerce')

    # 휴가시간이 08:00인 경우 시작시각을 "09:30"으로, 종료시각을 "18:30"으로 설정
    df.loc[df['휴가시간'] == '8:00', '시작시각'] = pd.to_datetime('09:30', format='%H:%M')
    df.loc[df['휴가시간'] == '8:00', '종료시각'] = pd.to_datetime('18:30', format='%H:%M')

    # 시작 시간이 08:00 이전인 경우 08:00으로 변경
    df.loc[df['시작시각'] < pd.to_datetime('08:00', format='%H:%M'), '시작시각'] = pd.to_datetime('08:00', format='%H:%M')

    # 동일 날짜에 대해 시작시간은 가장 이전 퇴근 시간을 가장 이후의 값을 합산
    df['시작시각'] = df.groupby(['이름', '날짜'])['시작시각'].transform('min')
    df['종료시각'] = df.groupby(['이름', '날짜'])['종료시각'].transform('max')

    # 중복 되는 값 삭제 처리
    df = df.drop_duplicates(subset=['이름', '날짜'])

    # 시작/종료 시간 컬럼이 비어있을 경우 휴일 또는 데이터 누락으로 생각하여 삭제 처리
    df = df.dropna(subset=['시작시각', '종료시각'], how='all')

    # 지각 여부 컬럼 추가
    df['지각'] = df['시작시각'].apply(lambda x: '지각' if x.time() >= pd.Timestamp('1900-01-01 10:30:00').time() else '정상')

    # 이름을 기준으로 근무 일수 설정 (여기서 일수 값이 이상하면 출/퇴근 체크를 누락한것)
    df_result = df.groupby('이름').size().reset_index(name='일수')

    # 지각 합산
    # 지각 여부에 따라 지각횟수 계산
    df_late = df.groupby('이름')['지각'].apply(lambda x: (x == '지각').sum()).reset_index(name='지각횟수')
    df_result = pd.merge(df_result, df_late, on='이름')

    # 기본 근무 시간 계산
    df['기본근무시간'] = df.apply(calculate_working_hours, axis=1)

    # 기본 근무 시간이 9시간을 넘어갈 경우 9시간으로 변경
    df['기본근무시간'] = df['기본근무시간'].apply(lambda x: timedelta(hours=9) if x > timedelta(hours=9) else x)

    df_weekly_time = df.groupby('이름')['기본근무시간'].sum()
    df_result = pd.merge(df_result, df_weekly_time, on='이름', how='inner')

    # 전체 근무 시간 계산
    df['총근무시간'] = df.apply(calculate_working_hours, axis=1)

    # 조직/직무 정보 추가
    df_cell = df.groupby('이름').agg({'조직': 'first', '역할(직무)': 'first'}).reset_index()
    df_result = pd.merge(df_result, df_cell, on='이름', how='inner')

    df_time = df.groupby('이름')['총근무시간'].sum()
    df_result = pd.merge(df_result, df_time, on='이름', how='inner')

    # 연장 근무 시간 계산
    df_result['근무시간'] = df_result['일수'].apply(lambda x: timedelta(hours=x * 8))

    df_result['연장'] = df_result['총근무시간'] - df_result['근무시간']
    df_result['연장'] = df_result['연장'].apply(lambda x: max(x, timedelta(hours=0)))

    # '기본근무시간' 컬럼 업데이트
    df_result.loc[df_result['기본근무시간'] > df_result['근무시간'], '기본근무시간'] = df_result['근무시간']

    # 시간 읽기 편하게 수정
    df_result['총근무시간'] = df_result['총근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df_result['기본근무시간'] = df_result['기본근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df_result['연장'] = df_result['연장'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    # 컬럼 순서 변경
    df_result = df_result[['이름', '조직', '역할(직무)', '일수', '총근무시간', '기본근무시간', '연장', '지각횟수']]

    # 가공 데이터
    keep = ['이름', '조직', '역할(직무)', '날짜', '시작시각', '종료시각', '지각', '기본근무시간', '총근무시간']
    df = df[keep]

    # 날짜 컬럼 값을 가져와서 적용
    df['시작시각'] = df['날짜'].dt.strftime('%Y-%m-%d') + ' ' + df['시작시각'].dt.strftime('%H:%M:%S')
    df['종료시각'] = df['날짜'].dt.strftime('%Y-%m-%d') + ' ' + df['종료시각'].dt.strftime('%H:%M:%S')

    # 날짜 데이터 yyyy/mm/dd 로 변환
    df['날짜'] = df['날짜'].dt.strftime('%Y-%m-%d')

    df['총근무시간'] = df['총근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df['기본근무시간'] = df['기본근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    return df_result, df


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5001)
    #app.run(port=5001)
