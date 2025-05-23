from flask import Flask, request, render_template, send_from_directory, redirect, url_for
from werkzeug.utils import secure_filename
import os
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
EXCEL_FILE = os.path.join(UPLOAD_FOLDER, 'du_lieu_nhap.xlsx')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(columns=[
        'Date', 'Số PBH', 'Tên HUB', 'Mã LK', 'Mô tả LK', 'Số lượng',
        'Serial máy', 'Serial tab', 'Serial module',
        'Ảnh Serial máy', 'Ảnh Serial tab', 'Ảnh Serial module'
    ])
    df_init.to_excel(EXCEL_FILE, index=False)

save_count = 0

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    global save_count
    data = request.form
    files = request.files

    def save_image(file):
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            return filename
        return ''

    img_serial_may = save_image(files.get('img_serial_may'))
    img_serial_tab = save_image(files.get('img_serial_tab'))
    img_serial_module = save_image(files.get('img_serial_module'))

    df = pd.read_excel(EXCEL_FILE)
    new_row = {
        'Date': data.get('date'),
        'Số PBH': data.get('pbh'),
        'Tên HUB': data.get('hub'),
        'Mã LK': data.get('ma_lk'),
        'Mô tả LK': data.get('mo_ta_lk'),
        'Số lượng': data.get('so_luong'),
        'Serial máy': data.get('serial_may'),
        'Serial tab': data.get('serial_tab'),
        'Serial module': data.get('serial_module'),
        'Ảnh Serial máy': img_serial_may,
        'Ảnh Serial tab': img_serial_tab,
        'Ảnh Serial module': img_serial_module
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    save_count += 1
    return redirect(url_for('index'))

@app.route('/downloads')
def downloads():
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'du_lieu_nhap.xlsx', as_attachment=True)

@app.context_processor
def inject_save_count():
    return dict(saveCount=save_count)

if __name__ == '__main__':
    app.run(debug=True)
