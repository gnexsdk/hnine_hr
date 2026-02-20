import os
import re
import logging
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
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from email.header import Header
from datetime import timedelta
from datetime import datetime

app = Flask(__name__)

# ============================================================
# 로깅 설정
# ============================================================
LOG_DIR = 'logs'
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 파일 핸들러 (일별 로그)
file_handler = logging.FileHandler(
    os.path.join(LOG_DIR, f'app_{datetime.now().strftime("%Y%m%d")}.log'),
    encoding='utf-8'
)
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(logging.Formatter(
    '[%(asctime)s] %(levelname)s [%(funcName)s:%(lineno)d] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
))

# 콘솔 핸들러
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter(
    '[%(asctime)s] %(levelname)s %(message)s',
    datefmt='%H:%M:%S'
))

# Flask 앱 로거 설정
app.logger.addHandler(file_handler)
app.logger.addHandler(console_handler)
app.logger.setLevel(logging.DEBUG)

# 모듈 로거 (Flask 외부 함수용)
logger = logging.getLogger(__name__)
logger.addHandler(file_handler)
logger.addHandler(console_handler)
logger.setLevel(logging.DEBUG)


# 결과 파일을 저장할 폴더 경로
RESULT_FOLDER = 'result'

# 기본 파일 경로 설정
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_RIP_DIR = BASE_DIR / 'static' / 'rip'
DEFAULT_RECIPIENT_FILE = DEFAULT_RIP_DIR / 'recv_list.xlsx'
DEFAULT_TEMPLATE_FILE = DEFAULT_RIP_DIR / 'mail_body.txt'


# ============================================================
# 요청 전후 로깅 미들웨어
# ============================================================
@app.before_request
def log_request_info():
    app.logger.info(f"[REQUEST] {request.method} {request.path} - IP: {request.remote_addr}")
    if request.method == 'POST':
        if request.files:
            for key, f in request.files.items():
                app.logger.info(f"  파일 업로드: {key} = {f.filename} ({f.content_type})")
        if request.form:
            safe_keys = ['team', 'name', 'relation', 'deceased', 'date',
                        'funeral_home', 'address', 'final_date', 'send']
            form_log = {k: v for k, v in request.form.items() if k in safe_keys}
            if form_log:
                app.logger.debug(f"  폼 데이터: {form_log}")


@app.after_request
def log_response_info(response):
    app.logger.info(f"[RESPONSE] {request.method} {request.path} -> {response.status_code}")
    return response


@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.error(f"[ERROR] {request.method} {request.path} - {type(e).__name__}: {str(e)}", exc_info=True)
    return f"서버 오류가 발생했습니다: {str(e)}", 500


# ============================================================
# 부고 이미지 생성
# ============================================================
def create_rip_image(team, name, relation, deceased, date, funeral_home, address, final_date):
    logger.info(f"부고 이미지 생성 시작: {team} {name}")

    width = 800
    height = 1000
    image = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(image)

    try:
        title_font = ImageFont.truetype("malgun.ttf", 30)
        body_font = ImageFont.truetype("malgun.ttf", 24)
        address_font = ImageFont.truetype("malgun.ttf", 22)
    except:
        try:
            title_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 30)
            body_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 24)
            address_font = ImageFont.truetype("C:\\Windows\\Fonts\\malgun.ttf", 22)
        except:
            try:
                title_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 30)
                body_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 24)
                address_font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 22)
            except:
                try:
                    title_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 30)
                    body_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 24)
                    address_font = ImageFont.truetype("/usr/share/fonts/nanum/NanumGothic.ttf", 22)
                except:
                    try:
                        title_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 30)
                        body_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 24)
                        address_font = ImageFont.truetype("/usr/share/fonts/NanumGothic.ttf", 22)
                    except:
                        try:
                            title_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 30)
                            body_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 24)
                            address_font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 22)
                        except Exception as e:
                            logger.error(f"폰트 로드 실패: {str(e)}")
                            raise Exception("사용 가능한 한글 폰트를 찾을 수 없습니다.")

    title = f"[訃告] {team} {name}님의 {relation}상"
    title_width = draw.textlength(title, font=title_font)
    title_x = (width - title_width) / 2
    draw.text((title_x, 50), title, font=title_font, fill='black')

    info_start_x = width * 0.2

    content_lines = [
        (f"{name}님의 {relation}(故 {deceased})께서", body_font, True),
        (f"{date}", body_font, True),
        ("별세하셨기에 삼가 알려 드립니다.", body_font, True),
        ("", body_font, True),
        (f"일  시 : {date}", body_font, False),
        (f"빈  소 : {funeral_home}", body_font, False),
        (f"({address})", address_font, False),
        (f"발  인 : {final_date}", body_font, False),
        ("", body_font, True),
        ("", body_font, True),
        ("삼가 고인의 명복을 빕니다.", body_font, True)
    ]

    y = 200
    for line, font, center_align in content_lines:
        if line.strip():
            if center_align:
                line_width = draw.textlength(line, font=font)
                x = (width - line_width) / 2
            else:
                x = info_start_x
                if line.startswith('('):
                    prefix_width = draw.textlength("빈  소 : ", font=body_font)
                    x = info_start_x + prefix_width
            draw.text((x, y), line, font=font, fill='black')
        y += 50

    img_byte_arr = BytesIO()
    image.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)

    logger.info("부고 이미지 생성 완료")
    return img_byte_arr


# ============================================================
# SMTP / 메일 발송
# ============================================================
def smtp_setting(email, encoded_password):
    logger.info(f"SMTP 연결 시작: {email}")
    password = base64.b64decode(encoded_password).decode('utf-8')

    port = 587
    mail_type = 'smtp.gmail.com'
    smtp = smtplib.SMTP(mail_type, port)
    smtp.set_debuglevel(True)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(email, password)

    logger.info("SMTP 연결 성공")
    return smtp


def send_rip_mail(sender, receiver, image_data, title_info, is_test=False):
    logger.info(f"부고 메일 발송 시작 - 수신자: {receiver}, 테스트: {is_test}")

    msg = MIMEMultipart('related')
    msg.set_charset('utf-8')

    mail_title = f"[訃告] {title_info['team']} {title_info['name']}님의 {title_info['relation']}상"
    msg['Subject'] = mail_title
    msg['From'] = sender
    msg['To'] = receiver if not is_test else 'hr@hnine.com'

    url_html = f"<div><a href='{title_info['url']}'>온라인 부고장</a></div>" if title_info.get('url') else ""

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    <body>
        <div>경영기획팀에서 안내드립니다.</div>
        <div>{title_info['team']} {title_info['name']}님의 {title_info['relation']}(故{title_info['deceased']})님께서 별세하셨기에 삼가 알려드립니다.</div>
        <br>
        {url_html}
        <div><img src="cid:rip_image"></div>
    </body>
    </html>
    """

    html_part = MIMEText(html, 'html', 'utf-8')
    msg.attach(html_part)

    image = MIMEImage(image_data.getvalue())
    image.add_header('Content-ID', '<rip_image>')
    image.add_header('Content-Disposition', 'inline')
    msg.attach(image)

    smtp = smtp_setting('hr@hnine.com', 'ZnR2endqZWJieWt3cHdjZw==')

    try:
        if is_test:
            smtp.sendmail(sender, 'hr@hnine.com', msg.as_string())
        else:
            smtp.sendmail(sender, receiver, msg.as_string())
        logger.info(f'메일 전송 완료: {receiver if not is_test else "hr@hnine.com"}')
    except Exception as e:
        logger.error(f"메일 전송 실패: {str(e)}", exc_info=True)
        raise e
    finally:
        smtp.quit()


# ============================================================
# 라우트: 부고 메일
# ============================================================
@app.route('/rip', methods=['GET', 'POST'])
def rip_mail():
    preview_image = None
    form_data = {
        'team': '', 'name': '', 'relation': '', 'deceased': '',
        'date': '', 'funeral_home': '', 'address': '', 'final_date': '', 'url': '',
    }

    if request.method == 'POST':
        try:
            form_data = {
                'team': request.form['team'],
                'name': request.form['name'],
                'relation': request.form['relation'],
                'deceased': request.form['deceased'],
                'date': request.form['date'],
                'funeral_home': request.form['funeral_home'],
                'address': request.form['address'],
                'final_date': request.form['final_date'],
                'url': request.form.get('url', '')
            }
            app.logger.info(f"부고 폼 제출: {form_data['team']} {form_data['name']}")

            image_data = create_rip_image(
                form_data['team'],
                form_data['name'],
                form_data['relation'],
                form_data['deceased'],
                form_data['date'],
                form_data['funeral_home'],
                form_data['address'],
                form_data['final_date'],
            )

            preview_image = base64.b64encode(image_data.getvalue()).decode('utf-8')

            if 'send' in request.form:
                is_test = request.form['send'] == 'test'
                title_info = {
                    'team': form_data['team'],
                    'name': form_data['name'],
                    'relation': form_data['relation'],
                    'deceased': form_data['deceased'],
                    'url': form_data['url']
                }
                send_rip_mail('hr@hnine.com', 'h9@hnine.com', image_data, title_info, is_test)
                app.logger.info(f"부고 메일 발송 완료 (테스트: {is_test})")
                return "메일 발송이 완료되었습니다." if not is_test else "테스트 메일이 발송되었습니다."

        except Exception as e:
            app.logger.error(f"부고 처리 중 오류: {str(e)}", exc_info=True)
            return f"오류 발생: {str(e)}"

    return render_template('rip.html', preview_image=preview_image, form_data=form_data, active_page='rip')


# ============================================================
# 야간 근무 처리
# ============================================================
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
        app.logger.info("야근 기록 파일 업로드 시작")

        work_file = request.files['file1']
        over_work_file = request.files['file2']

        app.logger.info(f"파일1: {work_file.filename}, 파일2: {over_work_file.filename}")

        df = pd.read_excel(work_file, sheet_name="result_row")
        df_over_work = pd.read_excel(over_work_file)

        app.logger.info(f"데이터 로드 완료 - 근무기록: {len(df)}행, 야근신청: {len(df_over_work)}행")

        df_original, df_result = process_overwork_xlsx(df, df_over_work)

        app.logger.info(f"야근 처리 완료 - 결과: {len(df_result)}행")

        if not os.path.exists(RESULT_FOLDER):
            os.makedirs(RESULT_FOLDER)

        original_filename = over_work_file.filename
        filename, extension = os.path.splitext(original_filename)
        result_filename = f"{filename}_result{extension}"

        result_filepath = os.path.join(RESULT_FOLDER, result_filename)

        with pd.ExcelWriter(result_filepath, engine='xlsxwriter') as writer:
            df_original.to_excel(writer, sheet_name='original', index=False)
            df_result.to_excel(writer, sheet_name='result', index=False)

        app.logger.info(f"결과 파일 저장: {result_filepath}")

        df_html = df_result.to_html()

        return render_template('overwork.html', df_html=df_html,
                             over_work_file_download=result_filepath, active_page='overwork')

    return render_template('overwork.html', active_page='overwork')


def process_overwork_xlsx(df, df_overwork):
    logger.info("야근 데이터 처리 시작")

    df_overwork['결과'] = ''
    df_overwork['날짜'] = pd.to_datetime(df_overwork['근무 일자']).dt.date

    df_overwork['결과'] = '확인필요'

    df_overwork['총근무시간'] = ''
    df_overwork = df_overwork[df_overwork["근무 유형"].apply(check_night_shift)]

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

    logger.info(f"야근 데이터 처리 완료: {len(df_result)}건")
    return df_overwork, df_result


# ============================================================
# 라우트: 근무기록 관리 (메인 페이지)
# ============================================================
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        app.logger.info(f"근무기록 파일 업로드: {uploaded_file.filename}")

        df = pd.read_excel(uploaded_file)
        app.logger.info(f"엑셀 데이터 로드 완료: {len(df)}행")

        df_result, df_result_row = process_xlsx(df)
        app.logger.info(f"근무기록 처리 완료 - 결과: {len(df_result)}명")

        if not os.path.exists(RESULT_FOLDER):
            os.makedirs(RESULT_FOLDER)

        original_filename = uploaded_file.filename
        filename, extension = os.path.splitext(original_filename)
        result_filename = f"{filename}_result{extension}"

        result_filepath = os.path.join(RESULT_FOLDER, result_filename)

        writer = pd.ExcelWriter(result_filepath, engine='xlsxwriter')

        df.to_excel(writer, sheet_name='Original', index=False)
        df_result.to_excel(writer, sheet_name='result', index=False)
        df_result_row.to_excel(writer, sheet_name='result_row', index=False)

        writer.close()

        app.logger.info(f"결과 파일 저장: {result_filepath}")

        df_html = df_result.to_html()

        return render_template('index.html', df_html=df_html,
                             file_download=result_filepath, active_page='worklog')

    return render_template('index.html', active_page='worklog')


@app.route('/download_overwork')
def download_overwork():
    output = request.args.get('over_work_file_download', None)
    app.logger.info(f"야근 결과 파일 다운로드: {output}")
    return send_file(output, as_attachment=True)


@app.route('/download_excel')
def download_excel():
    output = request.args.get('file_download', None)
    app.logger.info(f"근무기록 결과 파일 다운로드: {output}")
    return send_file(output, as_attachment=True)


# ============================================================
# 근무 시간 계산 함수
# ============================================================
def calculate_working_hours(row):
    if row['종료시각'] < row['시작시각']:
        working_hours = row['종료시각'] + timedelta(days=1) - row['시작시각'] - timedelta(hours=2)
    else:
        working_hours = row['종료시각'] - row['시작시각'] - timedelta(hours=1)
    return working_hours


def process_xlsx(df):
    logger.info(f"근무기록 데이터 처리 시작: {len(df)}행")

    df = df.dropna(subset=['조직'])
    df = df[~df['근무유형'].str.contains('수원')]

    df['날짜'] = pd.to_datetime(df['날짜'])
    df = df[~df['날짜'].dt.dayofweek.isin([5, 6])]

    logger.debug(f"주말 제거 후: {len(df)}행")

    df['시작시각'] = pd.to_datetime(df['시작시각'], format='%H:%M', errors='coerce')
    df['종료시각'] = pd.to_datetime(df['종료시각'], format='%H:%M', errors='coerce')

    df.loc[df['휴가시간'] == '8:00', '시작시각'] = pd.to_datetime('09:30', format='%H:%M')
    df.loc[df['휴가시간'] == '8:00', '종료시각'] = pd.to_datetime('18:30', format='%H:%M')

    df.loc[df['시작시각'] < pd.to_datetime('08:00', format='%H:%M'), '시작시각'] = pd.to_datetime('08:00', format='%H:%M')

    df['시작시각'] = df.groupby(['이름', '날짜'])['시작시각'].transform('min')
    df['종료시각'] = df.groupby(['이름', '날짜'])['종료시각'].transform('max')

    df = df.drop_duplicates(subset=['이름', '날짜'])
    df = df.dropna(subset=['시작시각', '종료시각'], how='all')

    df['지각'] = df['시작시각'].apply(lambda x: '지각' if x.time() >= pd.Timestamp('1900-01-01 10:30:00').time() else '정상')

    df_result = df.groupby('이름').size().reset_index(name='일수')

    df_late = df.groupby('이름')['지각'].apply(lambda x: (x == '지각').sum()).reset_index(name='지각횟수')
    df_result = pd.merge(df_result, df_late, on='이름')

    df['기본근무시간'] = df.apply(calculate_working_hours, axis=1)
    df['기본근무시간'] = df['기본근무시간'].apply(lambda x: timedelta(hours=9) if x > timedelta(hours=9) else x)

    df_weekly_time = df.groupby('이름')['기본근무시간'].sum()
    df_result = pd.merge(df_result, df_weekly_time, on='이름', how='inner')

    df['총근무시간'] = df.apply(calculate_working_hours, axis=1)

    df_cell = df.groupby('이름').agg({'조직': 'first', '역할(직무)': 'first'}).reset_index()
    df_result = pd.merge(df_result, df_cell, on='이름', how='inner')

    df_time = df.groupby('이름')['총근무시간'].sum()
    df_result = pd.merge(df_result, df_time, on='이름', how='inner')

    df_result['근무시간'] = df_result['일수'].apply(lambda x: timedelta(hours=x * 8))
    df_result['연장'] = df_result['총근무시간'] - df_result['근무시간']
    df_result['연장'] = df_result['연장'].apply(lambda x: max(x, timedelta(hours=0)))

    df_result.loc[df_result['기본근무시간'] > df_result['근무시간'], '기본근무시간'] = df_result['근무시간']

    df_result['총근무시간'] = df_result['총근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df_result['기본근무시간'] = df_result['기본근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df_result['연장'] = df_result['연장'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df_result = df_result[['이름', '조직', '역할(직무)', '일수', '총근무시간', '기본근무시간', '연장', '지각횟수']]

    keep = ['이름', '조직', '역할(직무)', '날짜', '시작시각', '종료시각', '지각', '기본근무시간', '총근무시간']
    df = df[keep]

    df['시작시각'] = df['날짜'].dt.strftime('%Y-%m-%d') + ' ' + df['시작시각'].dt.strftime('%H:%M:%S')
    df['종료시각'] = df['날짜'].dt.strftime('%Y-%m-%d') + ' ' + df['종료시각'].dt.strftime('%H:%M:%S')

    df['날짜'] = df['날짜'].dt.strftime('%Y-%m-%d')

    df['총근무시간'] = df['총근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    df['기본근무시간'] = df['기본근무시간'].apply(
        lambda x: f"{int(x.total_seconds() // 3600)}시간 {int((x.total_seconds() % 3600) // 60)}분")

    logger.info(f"근무기록 처리 완료: {len(df_result)}명")
    return df_result, df


if __name__ == '__main__':
    app.logger.info("=" * 60)
    app.logger.info("HINE Attendance Management 서버 시작")
    app.logger.info("=" * 60)
    app.run(host="0.0.0.0", port=5001)
    #app.run(port=5001)
