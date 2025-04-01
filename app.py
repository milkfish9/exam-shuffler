from flask import Flask, render_template, request, send_file
import pandas as pd
import random
from io import BytesIO
import zipfile

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/shuffle', methods=['POST'])
def shuffle():
    file = request.files['file']
    df = pd.read_excel(file)

    # 打亂題目順序
    df = df.sample(frac=1).reset_index(drop=True)

    questions = []
    answer_key = []

    for idx, row in df.iterrows():
        options = ['A選項', 'B選項', 'C選項', 'D選項']
        option_values = [row[opt] for opt in options]
        correct_answer = row['正解']

        option_pairs = list(zip(['A', 'B', 'C', 'D'], option_values))
        random.shuffle(option_pairs)

        new_correct_letter = ''
        new_options = {}

        for i, (label, value) in enumerate(option_pairs):
            new_label = ['A', 'B', 'C', 'D'][i]
            new_options[f'{new_label}選項'] = value
            if label == correct_answer:
                new_correct_letter = new_label

        questions.append({
            '題號': idx + 1,
            '題目': row['題目'],
            **new_options
        })

        answer_key.append({
            '題號': idx + 1,
            '正解': new_correct_letter
        })

    # 建立兩個 Excel 資料流
    exam_df = pd.DataFrame(questions)
    answer_df = pd.DataFrame(answer_key)

    exam_io = BytesIO()
    answer_io = BytesIO()

    exam_df.to_excel(exam_io, index=False)
    answer_df.to_excel(answer_io, index=False)

    exam_io.seek(0)
    answer_io.seek(0)

    # 壓縮成 zip
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        zip_file.writestr('考卷.xlsx', exam_io.getvalue())
        zip_file.writestr('答案.xlsx', answer_io.getvalue())

    zip_buffer.seek(0)

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name='考卷打包.zip',
        mimetype='application/zip'
    )

@app.route('/sample')
def download_sample():
    sample_data = {
        '題號': [1],
        '題目': ['哆啦？'],
        'A選項': ['A夢'],
        'B選項': ['B夢'],
        'C選項': ['C夢'],
        'D選項': ['D夢'],
        '正解': ['A']
    }
    df = pd.DataFrame(sample_data)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='範例題庫.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)