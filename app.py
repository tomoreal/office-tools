"""
XBRL to Excel Converter - Web Application

【プログラム構成】
このファイルは以下の機能ブロックで構成されています:

1. APPLICATION SETUP (1-11行)
   - Flask アプリケーション初期化
   - 環境変数設定
   - 将来の分割先: web/app.py または api/app.py

2. MAIN ROUTE (13-79行)
   - メインページの表示とファイル変換処理
   - convert_xbrl_to_excel.py の process_xbrl_zips を呼び出し
   - 将来の分割先: web/routes/converter.py

3. BOOKMARKLET ROUTES (81-87行)
   - ブックマークレット用ページ表示
   - 将来の分割先: web/routes/bookmarklets.py

4. TEMP CLEAR ROUTE (89-99行)
   - 一時ファイルクリア機能
   - 将来の分割先: web/routes/admin.py

5. LOCAL TESTING ENTRY POINT (101-103行)
   - ローカル開発用のエントリポイント
   - 将来の分割先: dev/run_local.py

【設計思想】
- convert_xbrl_to_excel.py の薄いラッパー
- 既存のコマンドライン機能をWebインターフェースとして提供
- 後方互換性を維持したまま、将来的にはRESTful APIとして分離可能

【依存関係】
- convert_xbrl_to_excel.py (Core Logic)
- templates/index.html, bookmarklets.html, csv_bookmarklets.html (Views)
- index.cgi (CGI Entry Point)
"""

import os
import tempfile
import urllib.parse
import shutil
import json

# ========================================================================
# APPLICATION SETUP
# ========================================================================
# 【将来の分割先】web/app.py または api/app.py

# LiteSpeedサーバー（コアサーバー等）でのマルチスレッド問題を回避
os.environ['OPENBLAS_NUM_THREADS'] = "1"
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

app = Flask(__name__)
app.secret_key = 'xbrl_to_excel_secret'

# EDINET APIキー（環境変数または直接指定）
EDINET_API_KEY = os.environ.get('EDINET_API_KEY', '6ea174edf112439da66798a6d863a95d')

# ========================================================================
# MAIN ROUTE - ファイルアップロードと変換処理
# ========================================================================
# 【将来の分割先】web/routes/converter.py

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Lazy imports to speed up CGI startup on GET requests
        from werkzeug.utils import secure_filename
        import convert_xbrl_to_excel
        
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

# ========================================================================
# EDINET API ROUTES - EDINET API連携機能
# ========================================================================
# 【将来の分割先】web/routes/edinet.py

@app.route('/api/edinet/search', methods=['POST'])
def edinet_search_company():
    """企業名でEDINET企業を検索"""
    from edinet_api import EdinetAPI

    try:
        data = request.get_json()
        company_name = data.get('company_name', '')

        if not company_name:
            return {'error': '企業名を入力してください'}, 400

        api = EdinetAPI(EDINET_API_KEY)
        results = api.search_company(company_name)

        return {'results': results}, 200

    except Exception as e:
        app.logger.error(f"Error in edinet_search_company: {e}")
        return {'error': str(e)}, 500


@app.route('/api/edinet/reports', methods=['POST'])
def edinet_get_reports():
    """指定企業の有価証券報告書一覧を取得"""
    from edinet_api import EdinetAPI

    try:
        data = request.get_json()
        edinet_code = data.get('edinet_code', '')
        start_date = data.get('start_date')
        end_date = data.get('end_date')

        if not edinet_code:
            return {'error': 'EDINETコードを指定してください'}, 400

        api = EdinetAPI(EDINET_API_KEY)
        results = api.get_securities_reports(edinet_code, start_date, end_date)

        return {'results': results}, 200

    except Exception as e:
        app.logger.error(f"Error in edinet_get_reports: {e}")
        return {'error': str(e)}, 500


@app.route('/api/edinet/convert', methods=['POST'])
def edinet_download_and_convert():
    """EDINET APIからXBRLをダウンロードしてExcelに変換"""
    from edinet_api import EdinetAPI
    import convert_xbrl_to_excel

    try:
        data = request.get_json()
        doc_ids = data.get('doc_ids', [])

        if not doc_ids:
            return {'error': '書類IDを指定してください'}, 400

        # 一時ディレクトリ作成
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        if not os.path.exists(base_temp_dir):
            os.makedirs(base_temp_dir, exist_ok=True)

        temp_dir = tempfile.mkdtemp(dir=base_temp_dir)

        # EDINET APIからXBRLをダウンロード
        api = EdinetAPI(EDINET_API_KEY)
        downloaded_paths = []

        for doc_id in doc_ids:
            file_path = os.path.join(temp_dir, f"{doc_id}.zip")
            if api.download_xbrl(doc_id, file_path):
                downloaded_paths.append(file_path)

        if not downloaded_paths:
            return {'error': 'XBRLファイルのダウンロードに失敗しました'}, 500

        # Excelに変換
        out_excel = convert_xbrl_to_excel.process_xbrl_zips(downloaded_paths, output_dir=temp_dir)

        if out_excel and os.path.exists(out_excel):
            # ファイルを送信
            filename = os.path.basename(out_excel)
            encoded_filename = urllib.parse.quote(filename)

            response = send_file(
                out_excel,
                as_attachment=True,
                download_name=filename
            )

            response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_filename}"
            return response
        else:
            return {'error': 'Excelファイルの生成に失敗しました'}, 500

    except Exception as e:
        app.logger.error(f"Error in edinet_download_and_convert: {e}")
        return {'error': str(e)}, 500


# ========================================================================
# BOOKMARKLET ROUTES - ブックマークレット用ページ
# ========================================================================
# 【将来の分割先】web/routes/bookmarklets.py

@app.route('/bookmarklets')
def bookmarklets():
    return render_template('bookmarklets.html')

@app.route('/csv_bookmarklets')
def csv_bookmarklets():
    return render_template('csv_bookmarklets.html')

@app.route('/csv_converter.html')
def csv_converter():
    """CSV変換ツールのページ"""
    return send_file('csv_converter.html')

# ========================================================================
# TEMP CLEAR ROUTE - 一時ファイルクリア
# ========================================================================
# 【将来の分割先】web/routes/admin.py

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

# ========================================================================
# LOCAL TESTING ENTRY POINT
# ========================================================================
# 【将来の分割先】dev/run_local.py

if __name__ == '__main__':
    # Run dynamically for local testing. In production, use Gunicorn e.g., gunicorn -w 4 app:app
    app.run(debug=True, port=8000, host='0.0.0.0')
