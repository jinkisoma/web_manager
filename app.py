import os
import sqlite3
import psycopg2
from psycopg2.extras import DictCursor
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, Response, jsonify
import pandas as pd
import io
from datetime import datetime
from urllib.parse import quote
from openpyxl.styles import PatternFill, Font, Alignment
import calendar

app = Flask(__name__)
app.secret_key = 'supersecretkey'

DATABASE_URL = os.environ.get('DATABASE_URL')
ATTACHMENT_DIR = "attachments"

# --- 거래처 및 작업 데이터 관리 ---
CLIENT_WORK_DATA = {
    '로지비': {
        '라벨작업': {'content': '단상자 바코드작업', 'price': 100},
        '포장작업': {'content': '선물세트 포장', 'price': 500}
    },

    '비플레인': {
        '소분작업': {'content': '샘플 소분', 'price': 50},
        '검수작업': {'content': '제품 외관 검수', 'price': 80}
    },


    '릴라이블': {
 '소분작업': {'content': '샘플 소분', 'price': 50},
  '검수작업': {'content': '제품 외관 검수', 'price': 80},
    '라벨작업': {'content': '단상자 바코드작업', 'price': 100},
        '포장작업': {'content': '선물세트 포장', 'price': 500}
    },

    '릴라이블(대성)': {
 '소분작업': {'content': '샘플 소분', 'price': 50},
  '검수작업': {'content': '제품 외관 검수', 'price': 80},
    '라벨작업': {'content': '단상자 바코드작업', 'price': 100},
        '포장작업': {'content': '선물세트 포장', 'price': 500}
    },

    '릴라이블(랩)': {
 '소분작업': {'content': '샘플 소분', 'price': 50},
  '검수작업': {'content': '제품 외관 검수', 'price': 80},
    '라벨작업': {'content': '단상자 바코드작업', 'price': 100},
        '포장작업': {'content': '아웃박스를 열여서 유관검수후에 다시 재포장해서 닫는작업', 'price': 1000}

    }
}

# [추가] 엑셀 헤더 한글 매핑
HEADER_MAP = {
    'id': '고유번호',
    'work_date': '작업일자',
    'client': '거래처',
    'author': '작성자',
    'product_code': '업체상품코드',
    'tracking_number': '송장번호',
    'work_type': '작업 구분',
    'content': '내용',
    'product_name': '상품명',
    'quantity': '작업수량',
    'box_quantity': '박스수량',
    'unit_price': '금액(단가)',
    'total_amount': '합계',
    'remarks': '비고',
    'confirmed': '확정여부'
}

# [추가] 확정 취소 비밀번호
CONFIRM_CANCEL_PASSWORD = "1234"

def get_db_connection():
    """데이터베이스 연결을 가져오는 함수"""
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=DictCursor)
    else:
        conn = sqlite3.connect('data.db')
        conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """데이터베이스 테이블을 초기화하는 함수"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if DATABASE_URL: # PostgreSQL
        cursor.execute('''
           CREATE TABLE IF NOT EXISTS user_data (
               id SERIAL PRIMARY KEY,
               work_date DATE NOT NULL,
               client TEXT NOT NULL,
               author TEXT,
               product_code TEXT,
               tracking_number TEXT,
               work_type TEXT,
               content TEXT,
               product_name TEXT,
               quantity INTEGER,
               box_quantity INTEGER,
               unit_price REAL,
               total_amount REAL,
               attachment TEXT,
               remarks TEXT,
               confirmed BOOLEAN DEFAULT FALSE NOT NULL
           )
        ''')
    else: # SQLite
        cursor.execute('''
           CREATE TABLE IF NOT EXISTS user_data (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               work_date DATE NOT NULL,
               client TEXT NOT NULL,
               author TEXT,
               product_code TEXT,
               tracking_number TEXT,
               work_type TEXT,
               content TEXT,
               product_name TEXT,
               quantity INTEGER,
               box_quantity INTEGER,
               unit_price REAL,
               total_amount REAL,
               attachment TEXT,
               remarks TEXT,
               confirmed INTEGER DEFAULT 0 NOT NULL
           )
        ''')

    conn.commit()
    cursor.close()
    conn.close()
    
    if not os.path.exists(ATTACHMENT_DIR):
        os.makedirs(ATTACHMENT_DIR)

def get_filtered_data(start_date, end_date, search_keyword):
    """[신규] 필터링된 데이터를 조회하는 공통 함수"""
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'
    
    # 엑셀에 필요한 컬럼만 명시적으로 선택
    base_query = "SELECT * FROM user_data"
    conditions = []
    params = []

    # [수정] 기간별 조회 로직
    if start_date:
        conditions.append(f"work_date >= {placeholder}")
        params.append(start_date)
    if end_date:
        conditions.append(f"work_date <= {placeholder}")
        params.append(end_date)

    if search_keyword:
        search_term = f"%{search_keyword}%"
        search_condition = f"""(client LIKE {placeholder} OR author LIKE {placeholder} 
                                OR product_name LIKE {placeholder} OR content LIKE {placeholder}
                                OR tracking_number LIKE {placeholder})"""
        conditions.append(search_condition)
        params.extend([search_term] * 5)

    if conditions:
        query = f"{base_query} WHERE {' AND '.join(conditions)} ORDER BY work_date ASC, id ASC"
        cursor.execute(query, tuple(params))
    else:
        cursor.execute(f"{base_query} ORDER BY work_date ASC, id ASC")

    rows = cursor.fetchall()
    # [추가] 컬럼 이름을 가져오는 로직
    columns = [description[0] for description in cursor.description]
    conn.close()
    # [수정] 데이터와 컬럼 이름을 함께 반환
    return rows, columns

@app.template_filter('number_format')
def number_format(value):
    """Jinja2 템플릿에서 사용할 숫자 서식 필터 (None -> 0, 천 단위 쉼표, 소수점 제거)"""
    if value is None:
        return 0
    # 소수점을 버리고 정수로 변환 후, 천 단위 쉼표 추가
    return f"{int(value):,}"

@app.route('/')
def index():
    """메인 페이지: 데이터 목록 및 월별/검색 결과 표시"""
    # [수정] 기간별 조회 파라미터 처리
    start_date_str = request.args.get('start_date')
    end_date_str = request.args.get('end_date')
    search_keyword = request.args.get('keyword', '')
    
    # [추가] 초기 접속 시, 기본값을 현재 월의 시작일과 종료일로 설정
    if not start_date_str and not end_date_str:
        today = datetime.today()
        # 현재 월의 첫째 날
        start_date_str = today.replace(day=1).strftime('%Y-%m-%d')
        # 현재 월의 마지막 날
        _, last_day = calendar.monthrange(today.year, today.month)
        end_date_str = today.replace(day=last_day).strftime('%Y-%m-%d')

    # [수정] 공통 함수를 사용하여 데이터 조회
    rows, _ = get_filtered_data(start_date_str, end_date_str, search_keyword)

    # [추가] 조회된 데이터의 건수 계산
    total_count = len(rows)
    confirmed_count = sum(1 for row in rows if row['confirmed'])
    unconfirmed_count = total_count - confirmed_count

    today_date = datetime.today().strftime('%Y-%m-%d')
    clients = list(CLIENT_WORK_DATA.keys())
    return render_template('index.html', users=rows, 
                           start_date=start_date_str, end_date=end_date_str, keyword=search_keyword,
                           today_date=today_date, clients=clients,
                           total_count=total_count, confirmed_count=confirmed_count, unconfirmed_count=unconfirmed_count)

@app.route('/api/work-items/<client_name>')
def get_work_items(client_name):
    work_items = CLIENT_WORK_DATA.get(client_name, {})
    return jsonify(work_items)

@app.route('/add', methods=['POST'])
def add_user():
    """데이터 추가"""
    work_date = request.form['work_date']
    client = request.form.get('client_select') if request.form.get('client_select') != 'direct' else request.form.get('client_direct', '')
    author = request.form['author']
    product_code = request.form['product_code']
    tracking_number = request.form['tracking_number']
    work_type = request.form.get('work_type_select') if request.form.get('work_type_select') != 'direct' else request.form.get('work_type_direct', '')
    content = request.form['content']
    product_name = request.form['product_name']
    quantity_str = request.form.get('quantity')
    box_quantity_str = request.form.get('box_quantity')
    unit_price_str = request.form.get('unit_price')
    remarks = request.form['remarks']

    if not work_date or not client or not author:
        flash('작업일자, 거래처, 작성자는 필수 항목입니다.', 'error')
        return redirect(url_for('index'))

    quantity = int(quantity_str) if quantity_str else 0
    unit_price = float(unit_price_str) if unit_price_str else 0.0
    total_amount = quantity * unit_price
    box_quantity = int(box_quantity_str) if box_quantity_str else None

    attachment_filename = ''
    if 'attachment' in request.files:
        file = request.files['attachment']
        if file.filename != '':
            attachment_filename = file.filename
            file.save(os.path.join(ATTACHMENT_DIR, attachment_filename))

    conn = get_db_connection()
    cursor = conn.cursor()
    
    placeholder = '%s' if DATABASE_URL else '?'
    query = f"""INSERT INTO user_data 
                (work_date, client, author, product_code, tracking_number, work_type, content, product_name, quantity, box_quantity, unit_price, total_amount, attachment, remarks) 
                VALUES ({", ".join([placeholder]*14)})"""

    cursor.execute(query, (
        work_date, client, author, product_code, tracking_number, work_type, content, product_name, 
        quantity, box_quantity, unit_price, total_amount, attachment_filename, remarks
    ))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 추가되었습니다.', 'success')
    return redirect(url_for('index'))

@app.route('/edit/<int:id>')
def edit_form(id):
    """수정 폼 페이지"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    placeholder = '%s' if DATABASE_URL else '?'
    query = f"SELECT * FROM user_data WHERE id = {placeholder}"

    cursor.execute(query, (id,))
    user = cursor.fetchone()
    conn.close()
    
    if user is None:
        flash('해당 데이터를 찾을 수 없습니다.', 'error')
        return redirect(url_for('index'))
    
    # [추가] 확정된 데이터는 수정 페이지에서 read-only로 처리하기 위한 플래그 전달
    is_readonly = user['confirmed']
    
    # [추가] 소유권 확인을 위한 현재 사용자 이름 전달 (JavaScript에서 사용)
    current_author = request.args.get('current_author', '')
    if not is_readonly and user['author'] != current_author:
        flash('본인이 작성한 데이터만 수정할 수 있습니다.', 'error')
        return redirect(url_for('index'))

    if is_readonly:
        flash('확정된 데이터는 수정할 수 없습니다.', 'error')
    
    clients = list(CLIENT_WORK_DATA.keys())
    return render_template('edit.html', user=user, clients=clients, is_readonly=is_readonly)

@app.route('/update/<int:id>', methods=['POST'])
def update_user(id):
    """데이터 수정"""
    work_date = request.form['work_date']
    client = request.form.get('client_select') if request.form.get('client_select') != 'direct' else request.form.get('client_direct', '')
    author = request.form['author']
    product_code = request.form['product_code']
    tracking_number = request.form['tracking_number']
    work_type = request.form.get('work_type_select') if request.form.get('work_type_select') != 'direct' else request.form.get('work_type_direct', '')
    content = request.form['content']
    product_name = request.form['product_name']
    quantity_str = request.form.get('quantity')
    box_quantity_str = request.form.get('box_quantity')
    unit_price_str = request.form.get('unit_price')
    remarks = request.form['remarks']

    # [수정] 소유권 및 확정 여부 확인
    current_author = request.form.get('current_author')
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'
    cursor.execute(f"SELECT author, confirmed FROM user_data WHERE id = {placeholder}", (id,))
    record = cursor.fetchone()
    if record and (record['confirmed'] or record['author'] != current_author):
        conn.close()
        flash('확정된 데이터는 수정할 수 없습니다.', 'error')
        return redirect(url_for('index'))
    
    # 연결을 닫았으므로, 이후 로직을 위해 다시 열어야 함
    conn.close()

    if not work_date or not client or not author:
        flash('작업일자, 거래처, 작성자는 필수 항목입니다.', 'error')
        return redirect(url_for('edit_form', id=id))

    quantity = int(quantity_str) if quantity_str else 0
    unit_price = float(unit_price_str) if unit_price_str else 0.0
    total_amount = quantity * unit_price
    box_quantity = int(box_quantity_str) if box_quantity_str else None

    attachment_filename = request.form.get('existing_attachment', '')
    delete_attachment = request.form.get('delete_attachment')

    if delete_attachment and attachment_filename:
        attachment_path = os.path.join(ATTACHMENT_DIR, attachment_filename)
        if os.path.exists(attachment_path):
            os.remove(attachment_path)
        attachment_filename = ''

    if 'attachment' in request.files:
        file = request.files['attachment']
        if file.filename != '':
            if attachment_filename and os.path.exists(os.path.join(ATTACHMENT_DIR, attachment_filename)):
                 os.remove(os.path.join(ATTACHMENT_DIR, attachment_filename))
            attachment_filename = file.filename
            file.save(os.path.join(ATTACHMENT_DIR, attachment_filename))

    conn = get_db_connection()
    cursor = conn.cursor()

    query = f"""UPDATE user_data SET 
                work_date={placeholder}, client={placeholder}, author={placeholder}, product_code={placeholder}, 
                tracking_number={placeholder}, work_type={placeholder}, content={placeholder}, product_name={placeholder}, 
                quantity={placeholder}, box_quantity={placeholder}, unit_price={placeholder}, total_amount={placeholder}, 
                attachment={placeholder}, remarks={placeholder} 
                WHERE id={placeholder}"""

    cursor.execute(query, (
        work_date, client, author, product_code, tracking_number, work_type, content, product_name,
        quantity, box_quantity, unit_price, total_amount, attachment_filename, remarks, id
    ))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 수정되었습니다.', 'success')
    return redirect(url_for('index'))

@app.route('/delete/<int:id>')
def delete_user(id):
    """데이터 및 첨부파일 삭제"""
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'

    # [수정] 소유권 및 확정 여부 확인
    current_author = request.args.get('current_author')
    cursor.execute(f"SELECT author, confirmed, attachment FROM user_data WHERE id = {placeholder}", (id,))
    record = cursor.fetchone()

    if record and (record['confirmed'] or record['author'] != current_author):
        conn.close()
        flash('본인이 작성했거나 확정되지 않은 데이터만 삭제할 수 있습니다.', 'error')
        return redirect(request.referrer or url_for('index'))
    
    if record and record['attachment']:
        attachment_path = os.path.join(ATTACHMENT_DIR, record['attachment'])
        if os.path.exists(attachment_path):
            os.remove(attachment_path)
            flash(f"첨부파일 '{record['attachment']}'이(가) 삭제되었습니다.", 'success')

    cursor.execute(f"DELETE FROM user_data WHERE id = {placeholder}", (id,))
    conn.commit()
    conn.close()
    
    flash('데이터가 삭제되었습니다.', 'success')
    return redirect(request.referrer or url_for('index'))

# [추가] 단일 건 확정 처리
@app.route('/confirm/<int:id>', methods=['POST'])
def confirm_user(id):
    # [추가] 소유권 확인
    current_author = request.form.get('current_author')
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'

    cursor.execute(f"SELECT author FROM user_data WHERE id = {placeholder}", (id,))
    record = cursor.fetchone()
    if record and record['author'] != current_author:
        conn.close()
        flash('본인이 작성한 데이터만 확정할 수 있습니다.', 'error')
        return redirect(request.referrer or url_for('index'))

    cursor.execute(f"UPDATE user_data SET confirmed = TRUE WHERE id = {placeholder} AND author = {placeholder}", (id, current_author))
    conn.commit()
    conn.close()
    flash(f'고유번호 {id} 데이터가 확정 처리되었습니다.', 'success')
    return redirect(request.referrer or url_for('index'))

# [추가] 전체 확정 처리
@app.route('/confirm_all', methods=['POST'])
def confirm_all():
    # [추가] 소유권 확인
    current_author = request.form.get('current_author')
    ids_to_confirm = request.form.getlist('confirm_ids')
    if not ids_to_confirm:
        flash('확정할 데이터가 없습니다.', 'error')
        return redirect(request.referrer or url_for('index'))

    if not current_author:
        flash('작성자 정보가 없어 전체 확정을 진행할 수 없습니다.', 'error')
        return redirect(request.referrer or url_for('index'))

    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'
    
    if DATABASE_URL:
        cursor.execute(f"UPDATE user_data SET confirmed = TRUE WHERE author = %s AND id = ANY(%s::int[])", (current_author, ids_to_confirm))
    else: # SQLite
        ids_tuple = tuple(int(id) for id in ids_to_confirm)
        query = f"UPDATE user_data SET confirmed = TRUE WHERE author = ? AND id IN ({', '.join(['?'] * len(ids_tuple))})"
        cursor.execute(query, (current_author, *ids_tuple))
        
    conn.commit()
    conn.close()
    flash(f'{len(ids_to_confirm)}개의 데이터가 전체 확정 처리되었습니다.', 'success')
    return redirect(request.referrer or url_for('index'))

# [추가] 확정 취소 처리
@app.route('/unconfirm/<int:id>', methods=['POST'])
def unconfirm_user(id):
    password = request.form.get('password')
    current_author = request.form.get('current_author')

    if password != CONFIRM_CANCEL_PASSWORD:
        flash('비밀번호가 일치하지 않습니다.', 'error')
        return redirect(request.referrer or url_for('index'))
    
    # [추가] 소유권 확인
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'

    cursor.execute(f"SELECT author FROM user_data WHERE id = {placeholder}", (id,))
    record = cursor.fetchone()
    if record and record['author'] != current_author:
        conn.close()
        flash('본인이 작성한 데이터만 확정 취소할 수 있습니다.', 'error')
        return redirect(request.referrer or url_for('index'))

    cursor.execute(f"UPDATE user_data SET confirmed = FALSE WHERE id = {placeholder} AND author = {placeholder}", (id, current_author))
    conn.commit()
    conn.close()
    flash(f'고유번호 {id} 데이터의 확정이 취소되었습니다.', 'success')
    return redirect(request.referrer or url_for('index'))

@app.route('/download_excel')
def download_excel():
    """현재 조회된 데이터를 엑셀 파일로 다운로드"""
    # [수정] 기간별 조회 파라미터 처리
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    search_keyword = request.args.get('keyword')

    # [수정] index와 동일한 공통 함수를 재사용하여 데이터 일관성 및 안정성 확보
    rows, columns = get_filtered_data(start_date, end_date, search_keyword)
    
    # 조회된 데이터가 없는 경우 빈 데이터프레임 생성
    if not rows:
        df = pd.DataFrame()
    else:
        # [수정] 데이터와 컬럼 이름을 함께 사용하여 DataFrame 생성
        df = pd.DataFrame(rows, columns=columns)
        # 엑셀에 불필요한 컬럼 제거 (선택 사항)
        if 'attachment' in df.columns:
            df = df.drop(columns=['attachment'])
        # [추가] 데이터프레임의 헤더를 한글로 변경
        df.rename(columns=HEADER_MAP, inplace=True)

        # [추가] 확정여부 값을 '확정'/'미확정' 텍스트로 변경
        df['확정여부'] = df['확정여부'].apply(lambda x: '확정' if x else '미확정')

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='정산데이터')

    # [추가] 엑셀 스타일링
    workbook = writer.book
    worksheet = writer.sheets['정산데이터']

    # 헤더 스타일 설정
    header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # 연한 블루
    header_font = Font(bold=True)
    header_alignment = Alignment(horizontal='center', vertical='center')

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # 숫자 서식 및 열 너비 자동 조절
    for idx, col in enumerate(df.columns, 1):
        column_letter = worksheet.cell(row=1, column=idx).column_letter
        if col in ['작업수량', '박스수량', '금액(단가)', '합계']:
            for cell in worksheet[column_letter][1:]:
                cell.number_format = '#,##0'
        
        # 열 너비 자동 조절
        max_length = 0
        for cell in worksheet[column_letter]:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    writer.close()
    output.seek(0)
    
    # [수정] 다운로드 시점의 '년-월'을 기준으로 동적 파일명 생성
    current_month_str = datetime.today().strftime('%Y-%m')
    filename = f"{current_month_str} 정산노트.xlsx"
    encoded_filename = quote(filename)

    return Response(output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f"attachment; filename*=UTF-8''{encoded_filename}"})


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """첨부파일 다운로드"""
    return send_from_directory(ATTACHMENT_DIR, filename)

# 애플리케이션 시작 시 DB 및 디렉토리 초기화
init_db()

if __name__ == '__main__':
    app.run(debug=True, port=8000)