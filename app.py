import os
import sqlite3
import psycopg2
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, Response, jsonify
import pandas as pd
import io
from datetime import datetime

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
    '릴라이블': {},
    '릴라이블(대성)': {},
    '릴라이블(랩)': {}
}

def get_db_connection():
    """데이터베이스 연결을 가져오는 함수"""
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL)
    else:
        conn = sqlite3.connect('data.db')
        conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """데이터베이스 테이블을 초기화하는 함수"""
    if not os.path.exists(ATTACHMENT_DIR):
        os.makedirs(ATTACHMENT_DIR)
    
    if not DATABASE_URL:
        conn = get_db_connection()
        # 송장번호(tracking_number) 필드 추가
        conn.execute('''
           CREATE TABLE IF NOT EXISTS user_data (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               work_date TEXT NOT NULL,
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
               remarks TEXT
           )
        ''')
        conn.commit()
        conn.close()

@app.route('/')
def index():
    """메인 페이지: 데이터 목록 및 월별/검색 결과 표시"""
    month_query = request.args.get('month', datetime.today().strftime('%Y-%m'))
    search_keyword = request.args.get('keyword', '')
    
    rows = []
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'
    
    base_query = "SELECT * FROM user_data"
    conditions = []
    params = []

    if month_query:
        # DB 호환성을 위해 날짜 포맷 함수 분기 처리
        if DATABASE_URL: # PostgreSQL
            conditions.append(f"to_char(work_date, 'YYYY-MM') = {placeholder}")
        else: # SQLite
            conditions.append(f"strftime('%Y-%m', work_date) = {placeholder}")
        params.append(month_query)

    if search_keyword:
        search_term = f"%{search_keyword}%"
        search_condition = f"""(client LIKE {placeholder} OR author LIKE {placeholder} 
                                OR product_name LIKE {placeholder} OR content LIKE {placeholder}
                                OR tracking_number LIKE {placeholder})"""
        conditions.append(search_condition)
        params.extend([search_term] * 5)

    if conditions:
        query = f"{base_query} WHERE {' AND '.join(conditions)} ORDER BY id DESC"
        cursor.execute(query, tuple(params))
    else:
        pass

    rows = cursor.fetchall()
    conn.close()

    today_date = datetime.today().strftime('%Y-%m-%d')
    clients = list(CLIENT_WORK_DATA.keys())
    return render_template('index.html', users=rows, keyword=search_keyword, 
                           today_date=today_date, clients=clients, current_month=month_query)

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

    if not work_date or not client:
        flash('작업일자와 거래처는 필수 항목입니다.', 'error')
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
        quantity if quantity_str else None, 
        box_quantity, 
        unit_price if unit_price_str else None, 
        total_amount, 
        attachment_filename, remarks
    ))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 추가되었습니다.', 'success')
    return redirect(url_for('index', month=work_date[:7]))

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
    
    clients = list(CLIENT_WORK_DATA.keys())
    return render_template('edit.html', user=user, clients=clients)

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

    if not work_date or not client:
        flash('작업일자와 거래처는 필수 항목입니다.', 'error')
        return redirect(url_for('edit_form', id=id))

    quantity = int(quantity_str) if quantity_str else 0
    unit_price = float(unit_price_str) if unit_price_str else 0.0
    total_amount = quantity * unit_price
    box_quantity = int(box_quantity_str) if box_quantity_str else None

    conn = get_db_connection()
    cursor = conn.cursor()

    placeholder = '%s' if DATABASE_URL else '?'
    query = f"""UPDATE user_data SET 
                work_date={placeholder}, client={placeholder}, author={placeholder}, product_code={placeholder}, 
                tracking_number={placeholder}, work_type={placeholder}, content={placeholder}, product_name={placeholder}, 
                quantity={placeholder}, box_quantity={placeholder}, unit_price={placeholder}, total_amount={placeholder}, 
                remarks={placeholder} 
                WHERE id={placeholder}"""

    cursor.execute(query, (
        work_date, client, author, product_code, tracking_number, work_type, content, product_name,
        quantity if quantity_str else None,
        box_quantity,
        unit_price if unit_price_str else None,
        total_amount,
        remarks,
        id
    ))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 수정되었습니다.', 'success')
    return redirect(url_for('index', month=work_date[:7]))

@app.route('/delete/<int:id>')
def delete_user(id):
    """데이터 및 첨부파일 삭제"""
    conn = get_db_connection()
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'

    # 삭제 전 첨부파일 이름 가져오기
    cursor.execute(f"SELECT attachment FROM user_data WHERE id = {placeholder}", (id,))
    record = cursor.fetchone()
    
    if record and record['attachment']:
        attachment_path = os.path.join(ATTACHMENT_DIR, record['attachment'])
        if os.path.exists(attachment_path):
            os.remove(attachment_path)
            flash(f"첨부파일 '{record['attachment']}'이(가) 삭제되었습니다.", 'success')

    # 데이터베이스에서 레코드 삭제
    cursor.execute(f"DELETE FROM user_data WHERE id = {placeholder}", (id,))
    conn.commit()
    conn.close()
    
    flash('데이터가 삭제되었습니다.', 'success')
    return redirect(request.referrer or url_for('index'))


@app.route('/download_excel')
def download_excel():
    """현재 조회된 데이터를 엑셀 파일로 다운로드"""
    month_query = request.args.get('month')
    search_keyword = request.args.get('keyword')

    conn = get_db_connection()
    conn.row_factory = None 
    cursor = conn.cursor()
    placeholder = '%s' if DATABASE_URL else '?'

    base_query = "SELECT id, work_date, client, author, product_code, tracking_number, work_type, content, product_name, quantity, box_quantity, unit_price, total_amount, remarks FROM user_data"
    conditions = []
    params = []

    if month_query:
        # DB 호환성을 위해 날짜 포맷 함수 분기 처리
        if DATABASE_URL: # PostgreSQL
            conditions.append(f"to_char(work_date, 'YYYY-MM') = {placeholder}")
        else: # SQLite
            conditions.append(f"strftime('%Y-%m', work_date) = {placeholder}")
        params.append(month_query)

    if search_keyword:
        search_term = f"%{search_keyword}%"
        search_condition = f"""(client LIKE {placeholder} OR author LIKE {placeholder} 
                                OR product_name LIKE {placeholder} OR content LIKE {placeholder}
                                OR tracking_number LIKE {placeholder})"""
        conditions.append(search_condition)
        params.extend([search_term] * 5)

    if conditions:
        query = f"{base_query} WHERE {' AND '.join(conditions)} ORDER BY id DESC"
        cursor.execute(query, tuple(params))
    else:
        query = f"{base_query} ORDER BY id DESC"
        cursor.execute(query)

    rows = cursor.fetchall()
    columns = [description[0] for description in cursor.description]
    conn.close()

    df = pd.DataFrame(rows, columns=columns)
    
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='정산데이터')
    writer.close()
    output.seek(0)

    return Response(output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': 'attachment;filename=정산데이터.xlsx'})


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """첨부파일 다운로드"""
    return send_from_directory(ATTACHMENT_DIR, filename)

if __name__ == '__main__':
    # init_db() # DB를 처음 생성할 때만 주석을 풀고 실행하세요.
    app.run(debug=True)
