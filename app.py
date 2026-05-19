from flask import Flask, request, send_file, render_template, jsonify
import io
import os
from pl_builder import build_combined_pl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html', months=MONTHS)


@app.route('/generate', methods=['POST'])
def generate():
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({'error': 'Two files are required.'}), 400

    file1 = request.files['file1']
    file2 = request.files['file2']
    month1 = request.form.get('month1', 'March')
    month2 = request.form.get('month2', 'April')
    year = request.form.get('year', '2026')

    if not file1.filename or not file2.filename:
        return jsonify({'error': 'Both files must be selected.'}), 400

    if not allowed_file(file1.filename) or not allowed_file(file2.filename):
        return jsonify({'error': 'Only .xlsx and .xls files are supported.'}), 400

    try:
        xlsx_bytes = build_combined_pl(
            mar_file=io.BytesIO(file1.read()),
            apr_file=io.BytesIO(file2.read()),
            mar_label=month1,
            apr_label=month2,
            year=year,
        )
    except Exception as e:
        return jsonify({'error': f'Failed to generate report: {str(e)}'}), 500

    filename = f"Combined_PL_{month1}_{month2}_{year}.xlsx"
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
