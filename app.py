from flask import Flask, render_template, request, send_file
import pandas as pd
import random
from io import BytesIO
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/shuffle', methods=['POST'])
def shuffle():
    file = request.files['file']
    df = pd.read_excel(file)

    questions = []

    for _, row in df.iterrows():
        options = ['A選項', 'B選項', 'C選項', 'D選項']
        option_values = [row[opt] for opt in options]
        correct_answer = row['正解']

        # 將選項與其標籤打包後打亂
        option_pairs = list(zip(['A', 'B', 'C', 'D'], option_values))
        random.shuffle(option_pairs)

        new_correct_letter = ''
        new_options = {}

        for idx, (label, value) in enumerate(option_pairs):
            new_label = ['A', 'B', 'C', 'D'][idx]
            new_options[f'{new_label}選項'] = value
            if label == correct_answer:
                new_correct_letter = new_label

        questions.append({
            '題號': row['題號'],
            '題目': row['題目'],
            **new_options,
            '正解': new_correct_letter
        })

    shuffled_df = pd.DataFrame(questions)

    # 產生考卷與答案卷
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        shuffled_df.drop(columns=['正解']).to_excel(writer, index=False, sheet_name='題卷')
        shuffled_df[['題號', '正解']].to_excel(writer, index=False, sheet_name='答案卷')
    output.seek(0)

    return send_file(output, as_attachment=True, download_name='shuffled_exam.zip', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/sample')
def download_sample():
    sample_df = pd.DataFrame([{
        '題號': 1,
        '題目': '哆啦？',
        'A選項': 'A夢',
        'B選項': 'B夢',
        'C選項': 'C夢',
        'D選項': 'D夢',
        '正解': 'A'
    }])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='sample.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ✅ 這是對外開放、支援 Render 所需的啟動設定
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
