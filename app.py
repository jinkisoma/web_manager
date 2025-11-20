import os
import sqlite3
import psycopg2
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, Response
import pandas as pd # 엑셀 생성을 위해 pandas 임포트
import io # 메모리 상에서 엑셀 파일을 다루기 위해 io 임포트

app = Flask(__name__)
# flash 메시지를 사용하기 위한 시크릿 키 설정
app.secret_key = 'supersecretkey' 

# --- 데이터베이스 설정 ---
# Render 배포 환경에서는 DATABASE_URL 환경 변수를 사용하고,
# 로컬 환경에서는 sqlite3를 사용합니다.
DATABASE_URL = os.environ.get('DATABASE_URL')
ATTACHMENT_DIR = "attachments"

def get_db_connection():
    """데이터베이스 연결을 가져오는 함수"""
    if DATABASE_URL: # 배포 환경 (PostgreSQL)
        conn = psycopg2.connect(DATABASE_URL)
    else: # 로컬 환경 (SQLite)
        conn = sqlite3.connect('data.db')
        conn.row_factory = sqlite3.Row # 컬럼 이름으로 접근 가능하게 설정
    return conn

def init_db():
    """데이터베이스 테이블을 초기화하는 함수"""
    if not os.path.exists(ATTACHMENT_DIR):
        os.makedirs(ATTACHMENT_DIR)
        
    # 로컬 환경에서만 SQLite DB 파일 생성
    if not DATABASE_URL:
        conn = get_db_connection()
        # PostgreSQL과 호환되도록 id를 SERIAL PRIMARY KEY 처럼 동작하게 수정
        # TEXT 타입은 두 DB 모두에서 잘 동작함
        conn.execute('''
           CREATE TABLE IF NOT EXISTS user_data (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               name TEXT NOT NULL,
               phone TEXT,
               email TEXT,
               memo TEXT,
               attachment TEXT
           )
        ''')
        conn.commit()
        conn.close()

@app.route('/')
def index():
    """메인 페이지: 검색하거나 '전체보기'를 눌렀을 때만 데이터 표시"""
    search_keyword = request.args.get('keyword', '')
    show_all = request.args.get('show_all', 'false') # '전체보기' 플래그 추가
    
    rows = []
    
    if search_keyword or show_all == 'true':
        conn = get_db_connection()
        cursor = conn.cursor()
        placeholder = '%s' if DATABASE_URL else '?'

        if search_keyword: # 검색어가 있는 경우
            search_term = f"%{search_keyword}%"
            query = f"SELECT * FROM user_data WHERE name LIKE {placeholder} OR phone LIKE {placeholder} OR email LIKE {placeholder} OR memo LIKE {placeholder} ORDER BY id DESC"
            cursor.execute(query, (search_term, search_term, search_term, search_term))
        else: # '전체보기'가 요청된 경우
            cursor.execute("SELECT * FROM user_data ORDER BY id DESC")
            
        rows = cursor.fetchall()
        conn.close()

    return render_template('index.html', users=rows, keyword=search_keyword)

@app.route('/download_excel')
def download_excel():
    """현재 조회된 데이터를 엑셀 파일로 다운로드하는 기능"""
    search_keyword = request.args.get('keyword', '')

    conn = get_db_connection()
    cursor = conn.cursor()
    
    placeholder = '%s' if DATABASE_URL else '?'

    if search_keyword:
        query = f"SELECT id, name, phone, email, memo FROM user_data WHERE name LIKE {placeholder} OR phone LIKE {placeholder} OR email LIKE {placeholder} OR memo LIKE {placeholder} ORDER BY id DESC"
        search_term = f"%{search_keyword}%"
        cursor.execute(query, (search_term, search_term, search_term, search_term))
    else:
        query = "SELECT id, name, phone, email, memo FROM user_data ORDER BY id DESC"
        cursor.execute(query)
        
    rows = cursor.fetchall()
    conn.close()

    # pandas DataFrame으로 데이터 변환
    # fetchall() 결과가 DB 드라이버에 따라 튜플 리스트일 수 있으므로 컬럼명 지정
    df = pd.DataFrame(rows, columns=['id', 'name', 'phone', 'email', 'memo'])
    df.columns = ['ID', '이름', '연락처', '이메일', '메모'] # 엑셀에 표시될 컬럼명 변경

    # 메모리 버퍼를 사용하여 엑셀 파일 생성
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='데이터')
    writer.close()
    output.seek(0)

    # Flask Response로 파일 다운로드 응답 생성
    return Response(output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': 'attachment;filename=user_data.xlsx'})

@app.route('/add', methods=['POST'])
def add_user():
    """데이터 추가"""
    name = request.form['name']
    phone = request.form['phone']
    email = request.form['email']
    memo = request.form['memo']
    
    if not name:
        flash('이름은 필수 항목입니다.', 'error')
        return redirect(url_for('index'))

    attachment_filename = ''
    if 'attachment' in request.files:
        file = request.files['attachment']
        if file.filename != '':
            attachment_filename = file.filename
            file.save(os.path.join(ATTACHMENT_DIR, attachment_filename))

    conn = get_db_connection()
    cursor = conn.cursor()
    
    # DB 종류에 따라 다른 플레이스홀더 사용
    placeholder = '%s' if DATABASE_URL else '?'
    query = f"INSERT INTO user_data (name, phone, email, memo, attachment) VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})"

    cursor.execute(query, (name, phone, email, memo, attachment_filename))
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
    return render_template('edit.html', user=user)

@app.route('/update/<int:id>', methods=['POST'])
def update_user(id):
    """데이터 수정"""
    name = request.form['name']
    phone = request.form['phone']
    email = request.form['email']
    memo = request.form['memo']

    conn = get_db_connection()
    cursor = conn.cursor()

    placeholder = '%s' if DATABASE_URL else '?'
    query = f"UPDATE user_data SET name = {placeholder}, phone = {placeholder}, email = {placeholder}, memo = {placeholder} WHERE id = {placeholder}"

    cursor.execute(query, (name, phone, email, memo, id))
    conn.commit()
    conn.close()
    flash('데이터가 성공적으로 수정되었습니다.', 'success')
    return redirect(url_for('index'))

@app.route('/delete/<int:id>')
def delete_user(id):
    """데이터 삭제"""
    conn = get_db_connection()
    cursor = conn.cursor()

    placeholder = '%s' if DATABASE_URL else '?'
    query = f"DELETE FROM user_data WHERE id = {placeholder}"

    cursor.execute(query, (id,))
    conn.commit()
    conn.close()
    flash('데이터가 삭제되었습니다.', 'success')
    return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """첨부파일 다운로드"""
    return send_from_directory(ATTACHMENT_DIR, filename)

if __name__ == '__main__':
    init_db() # 로컬에서 실행 시 DB 초기화
    app.run(debug=True) # 개발용 서버 실행

