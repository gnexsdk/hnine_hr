import os
import smtplib
import pandas as pd
import base64
from pathlib import Path
from flask import Flask, request, render_template, send_file
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from PIL import Image, ImageDraw, ImageFont  # 이 줄을 추가
from io import BytesIO
from email.header import Header
from datetime import timedelta
from datetime import datetime

app = Flask(__name__)

# 결과 파일을 저장할 폴더 경로
RESULT_FOLDER = 'result'

# 기본 파일 경로 설정
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_RIP_DIR = BASE_DIR / 'static' / 'rip'
DEFAULT_RECIPIENT_FILE = DEFAULT_RIP_DIR / 'recv_list.xlsx'
DEFAULT_TEMPLATE_FILE = DEFAULT_RIP_DIR / 'mail_body.txt'


def create_rip_image(team, name, relation, deceased, date, funeral_home, address, final_date):
   # 흰색 배경의 이미지 생성
   width = 800
   height = 1000
   image = Image.new('RGB', (width, height), 'white')
   draw = ImageDraw.Draw(image)

   try:
       title_font = ImageFont.truetype("malgun.ttf", 30)
       body_font = ImageFont.truetype("malgun.ttf", 24)
       address_font = ImageFont.truetype("malgun.ttf", 22)  # 주소용 폰트 크기 작게
   except:
       try:
           # Windows의 맑은 고딕
           title_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 30)
           body_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 24)
           address_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 22)
       except:
           try:
               # Linux 일반적인 폰트 경로
               title_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 30)
               body_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 24)
               address_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 22)
           except:
               try:
                   # Ubuntu/Debian 계열 폰트 경로
                   title_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 30)
                   body_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 24)
                   address_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 22)
               except:
                   try:
                       # 다른 Linux 폰트 경로
                       title_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 30)
                       body_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 24)
                       address_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 22)
                   except:
                       try:
                           # macOS의 기본 한글 폰트
                           title_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 30)
                           body_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 24)
                           address_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 22)
                       except Exception as e:
                           print(f"폰트 로드 실패: {str(e)}")
                           raise Exception("사용 가능한 한글 폰트를 찾을 수 없습니다.")

   # 제목 작성 및 가운데 정렬
   title = f"[訃告] {team} {name}님의 {relation}상"
   title_width = draw.textlength(title, font=title_font)
   title_x = (width - title_width) / 2
   draw.text((title_x, 50), title, font=title_font, fill='black')

   # 정보 라인의 시작 위치 계산
   info_start_x = width * 0.2  # 왼쪽에서 20% 지점에서 시작

   # 본문 작성
   content_lines = [
       (f"{name}님의 {relation}(故 {deceased})께서", body_font, True),
       (f"{date}", body_font, True),
       ("별세하셨기에 삼가 알려 드립니다.", body_font, True),
       ("", body_font, True),
       (f"일  시 : {date}", body_font, False),
       (f"빈  소 : {funeral_home}", body_font, False),
       (f"({address})", address_font, False),  # 주소는 괄호로 감싸고 빈소와 같은 위치에서 시작
       (f"발  인 : {final_date}", body_font, False),
       ("", body_font, True),
       ("", body_font, True),
       ("삼가 고인의 명복을 빕니다.", body_font, True)
   ]

   # 각 줄을 그리기
   y = 200
   for line, font, center_align in content_lines:
       if line.strip():  # 빈 줄이 아닌 경우
           if center_align:
               # 가운데 정렬
               line_width = draw.textlength(line, font=font)
               x = (width - line_width) / 2
           else:
               # 일정한 위치에서 시작
               x = info_start_x
               if line.startswith('('):  # 주소인 경우
                   # 빈소의 "빈  소 : " 부분 너비 계산
                   prefix_width = draw.textlength("빈  소 : ", font=body_font)
                   x = info_start_x + prefix_width
           draw.text((x, y), line, font=font, fill='black')
       y += 50  # 줄 간격

   # 이미지를 바이트로 변환
   img_byte_arr = BytesIO()
   image.save(img_byte_arr, format='PNG')
   img_byte_arr.seek(0)

   return img_byte_arr


def smtp_setting(email, encoded_password):
    password = base64.b64decode(encoded_password).decode('utf-8')

    port = 587
    mail_type = 'smtp.gmail.com'
    smtp = smtplib.SMTP(mail_type, port)
    smtp.set_debuglevel(True)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(email, password)

    return smtp


def send_rip_mail(sender, receiver, image_data, title_info, is_test=False):
    msg = MIMEMultipart('related')  # related로 설정
    msg.set_charset('utf-8')

    # 제목 생성
    mail_title = f"[訃告] {title_info['team']} {title_info['name']}님의 {title_info['relation']}상"
    msg['Subject'] = mail_title
    msg['From'] = sender
    msg['To'] = receiver if not is_test else 'hr@hnine.com'

    # HTML 본문 생성 - 이미지를 본문에 삽입
    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    <body>
        <div>경영기획팀에서 안내드립니다.</div>
        <div>{title_info['team']} {title_info['name']}님의 {title_info['relation']}(故{title_info['deceased']})님께서 별세하셨기에 삼가 알려드립니다.</div>
        <div><img src="cid:rip_image"></div>
    </body>
    </html>
    """

    # HTML 파트 추가
    html_part = MIMEText(html, 'html', 'utf-8')
    msg.attach(html_part)

    # 이미지를 본문에 삽입하기 위한 설정
    image = MIMEImage(image_data.getvalue())
    image.add_header('Content-ID', '<rip_image>')
    image.add_header('Content-Disposition', 'inline')
    msg.attach(image)

    # SMTP 설정 및 발송
    smtp = smtp_setting('hr@hnine.com', 'ZnR2endqZWJieWt3cHdjZw==')

    try:
        if is_test:
            smtp.sendmail(sender, 'hr@hnine.com', msg.as_string())
        else:
            smtp.sendmail(sender, receiver, msg.as_string())
        print(f'메일 전송 완료: {receiver if not is_test else "hr@hnine.com"}')
    except Exception as e:
        print(f"Failed to send email: {str(e)}")
        raise e
    finally:
        smtp.quit()


@app.route('/rip', methods=['GET', 'POST'])
def rip_mail():
    preview_image = None
    # 폼 데이터를 저장할 변수들 초기화
    form_data = {
        'team': '',
        'name': '',
        'relation': '',
        'deceased': '',
        'date': '',
        'funeral_home': '',
        'address': '',
        'final_date': ''
    }

    if request.method == 'POST':
        try:
            # 폼 데이터 가져오기
            form_data = {
                'team': request.form['team'],
                'name': request.form['name'],
                'relation': request.form['relation'],
                'deceased': request.form['deceased'],
                'date': request.form['date'],
                'funeral_home': request.form['funeral_home'],
                'address': request.form['address'],
                'final_date': request.form['final_date']
            }

            # 이미지 생성
            image_data = create_rip_image(
                form_data['team'],
                form_data['name'],
                form_data['relation'],
                form_data['deceased'],
                form_data['date'],
                form_data['funeral_home'],
                form_data['address'],
                form_data['final_date']
            )

            # 이미지를 base64로 인코딩하여 HTML에서 표시
            preview_image = base64.b64encode(image_data.getvalue()).decode('utf-8')

            # 메일 발송 버튼이 클릭된 경우
            if 'send' in request.form:
                is_test = request.form['send'] == 'test'
                # 제목 정보 전달
                title_info = {
                    'team': form_data['team'],
                    'name': form_data['name'],
                    'relation': form_data['relation'],
                    'deceased': form_data['deceased']
                }
                send_rip_mail('hr@hnine.com', 'h9@hnine.com', image_data, title_info, is_test)
                return "메일 발송이 완료되었습니다." if not is_test else "테스트 메일이 발송되었습니다."

        except Exception as e:
            return f"오류 발생: {str(e)}"

    return render_template('rip.html', preview_image=preview_image, form_data=form_data)


def check_night_shift(x):
    if isinstance(x, list):
        return "야간 근무" in x
    elif isinstance(x, str):
        return "야간 근무" in x
    else:
        return False


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
    df_overwork['날짜'] = pd.to_datetime(df_overwork['근무 일자']).dt.date

    df_overwork['결과'] = '확인필요'

    df_overwork['총근무시간'] = ''
    df_overwork = df_overwork[df_overwork["근무 유형"].apply(check_night_shift)]

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
    df_result['야간 근무 일자'] = df_overwork['근무 일자']
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
