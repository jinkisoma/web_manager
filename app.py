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

# --- 1, 2. 거래처 및 거래처별 작업 데이터 관리 ---
# 새로운 거래처나 작업 내용을 추가하려면 이 딕셔너리를 수정하세요.
# '거래처명': {
#     '작업구분명': {'content': '상세내용', 'price': 단가},
#     ...
# }
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
        # 3. 새로운 입력폼에 맞게 DB 스키마 변경
        conn.execute('''
           CREATE TABLE IF NOT EXISTS user_data (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               work_date TEXT NOT NULL,
               client TEXT NOT NULL,
               author TEXT,
               product_code TEXT,
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
    """메인 페이지: 데이터 목록 및 검색 결과 표시"""
    search_keyword = request.args.get('keyword', '')
    show_all = request.args.get('show_all', 'false')
    
    rows = []
    
    if search_keyword or show_all == 'true':
        conn = get_db_connection()
        cursor = conn.cursor()
        placeholder = '%s' if DATABASE_URL else '?'
        
        # 검색 필드를 새로운 스키마에 맞게 수정
        if search_keyword:
            search_term = f"%{search_keyword}%"
            query = f"""SELECT * FROM user_data 
                        WHERE client LIKE {placeholder} OR author LIKE {placeholder} 
                        OR product_name LIKE {placeholder} OR content LIKE {placeholder} 
                        ORDER BY id DESC"""
            cursor.execute(query, (search_term, search_term, search_term, search_term))
        else:
            cursor.execute("SELECT * FROM user_data ORDER BY id DESC")
            
        rows = cursor.fetchall()
        conn.close()

    # 1. 오늘 날짜와 거래처 목록을 템플릿에 전달
    today_date = datetime.today().strftime('%Y-%m-%d')
    clients = list(CLIENT_WORK_DATA.keys())
    return render_template('index.html', users=rows, keyword=search_keyword, today_date=today_date, clients=clients)

# 4, 5. 거래처 선택 시 작업 구분을 동적으로 제공하는 API
@app.route('/api/work-items/<client_name>')
def get_work_items(client_name):
    work_items = CLIENT_WORK_DATA.get(client_name, {})
    return jsonify(work_items)

@app.route('/add', methods=['POST'])
def add_user():
    """데이터 추가"""
    # 3. 새로운 폼 필드에서 데이터 받아오기
    work_date = request.form['work_date']
    client = request.form.get('client_select') if request.form.get('client_select') != 'direct' else request.form.get('client_direct', '')
    author = request.form['author']
    product_code = request.form['product_code']
    work_type = request.form.get('work_type_select') if request.form.get('work_type_select') != 'direct' else request.form.get('work_type_direct', '')
    content = request.form['content']
    product_name = request.form['product_name']
    quantity = request.form.get('quantity')
    box_quantity = request.form.get('box_quantity')
    unit_price = request.form.get('unit_price')
    total_amount = request.form.get('total_amount')
    remarks = request.form['remarks']

    if not work_date or not client:
        flash('작업일자와 거래처는 필수 항목입니다.', 'error')
        return redirect(url_for('index'))

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
                (work_date, client, author, product_code, work_type, content, product_name, quantity, box_quantity, unit_price, total_amount, attachment, remarks) 
                VALUES ({", ".join([placeholder]*13)})"""

    # DB에 저장 (값이 비어있을 경우 None으로 변환)
    cursor.execute(query, (
        work_date, client, author, product_code, work_type, content, product_name, 
        int(quantity) if quantity else None, 
        int(box_quantity) if box_quantity else None, 
        float(unit_price) if unit_price else None, 
        float(total_amount) if total_amount else None, 
        attachment_filename, remarks
    ))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 추가되었습니다.', 'success')
    return redirect(url_for('index'))

# --- 기존 기능 (엑셀 다운로드, 수정, 삭제 등)은 새로운 스키마에 맞게 수정이 필요합니다. ---
# --- 현재는 주요 요청 기능 구현에 집중했습니다. ---

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """첨부파일 다운로드"""
    return send_from_directory(ATTACHMENT_DIR, filename)

if __name__ == '__main__':
    # init_db() # 애플리케이션 시작 시마다 DB를 초기화하지 않도록 주석 처리합니다.
    # 데이터베이스 테이블을 처음 생성하거나 구조를 변경할 때만 이 함수의 주석을 해제하고 한 번 실행하세요.
    app.run(debug=True)
