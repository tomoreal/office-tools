import os
import tempfile
import urllib.parse
import shutil

# LiteSpeedサーバー（コアサーバー等）でのマルチスレッド問題を回避
os.environ['OPENBLAS_NUM_THREADS'] = "1"
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

import convert_xbrl_to_excel
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'xbrl_to_excel_secret'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'files' not in request.files:
            flash('ファイルがアップロードされていません。')
            return redirect(request.url)
            
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            flash('ファイルが選択されていません。')
            return redirect(request.url)

        # プロジェクト内の temp_uploads ディレクトリを使用（権限問題を回避）
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        if not os.path.exists(base_temp_dir):
            os.makedirs(base_temp_dir, exist_ok=True)
            
        temp_dir = tempfile.mkdtemp(dir=base_temp_dir)
        
        saved_paths = []
        try:
            for file in files:
                if file and file.filename.endswith('.zip'):
                    filename = secure_filename(file.filename)
                    # if the user uploaded something with Japanese characters, secure_filename might empty it
                    # fallback if empty
                    if not filename:
                        filename = "upload.zip"
                    file_path = os.path.join(temp_dir, filename)
                    file.save(file_path)
                    saved_paths.append(file_path)
            
            if not saved_paths:
                flash('有効な .zip ファイルをアップロードしてください。')
                return redirect(request.url)
                
            # Call the updated parsing logic
            out_excel = convert_xbrl_to_excel.process_xbrl_zips(saved_paths, output_dir=temp_dir)
            
            if out_excel and os.path.exists(out_excel):
                # Send the Excel file back to the browser
                filename = os.path.basename(out_excel)
                encoded_filename = urllib.parse.quote(filename)
                
                response = send_file(
                    out_excel,
                    as_attachment=True,
                    download_name=filename
                )
                
                # Make sure the Japanese filename displays correctly in the browser download prompt
                response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_filename}"
                return response
            else:
                flash("Excelファイルの生成に失敗しました。")
                return redirect(request.url)
                
        except Exception as e:
            app.logger.error(f"Error during conversion: {e}")
            flash(f"エラーが発生しました: {str(e)}")
            return redirect(request.url)
            
    return render_template('index.html')
    
@app.route('/clear', methods=['POST'])
def clear_temp():
    try:
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        if os.path.exists(base_temp_dir):
            shutil.rmtree(base_temp_dir)
            os.makedirs(base_temp_dir, exist_ok=True)
        flash('サーバー上の一次ファイルをクリアしました。')
    except Exception as e:
        flash(f'クリア中にエラーが発生しました: {str(e)}')
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Run dynamically for local testing. In production, use Gunicorn e.g., gunicorn -w 4 app:app
    app.run(debug=True, port=8000, host='0.0.0.0')
