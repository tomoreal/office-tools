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

# EDINET APIキー
from edinet_api_key import EDINET_API_KEY

# プロジェクト内の temp_uploads ディレクトリを使用（権限問題を回避）
BASE_TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
if not os.path.exists(BASE_TEMP_DIR):
    os.makedirs(BASE_TEMP_DIR, exist_ok=True)

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
                # Return JSON with the relative file path for the download API
                relative_path = os.path.relpath(out_excel, BASE_TEMP_DIR)
                return {
                    "success": True,
                    "file": relative_path
                }
            else:
                return {
                    "success": False,
                    "error": "Excelファイルの生成に失敗しました。"
                }
                
        except Exception as e:
            app.logger.error(f"Error during conversion: {e}")
            return {
                "success": False,
                "error": str(e)
            }
            
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


@app.route('/api/edinet/preload', methods=['POST'])
def edinet_preload():
    """EDINET APIからXBRLを先読みダウンロード（並行処理・最大5並列・リトライ機能付き）"""
    from edinet_api import EdinetAPI
    from flask import jsonify
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time

    try:
        data = request.get_json()
        doc_ids = data.get('doc_ids', [])

        if not doc_ids:
            return jsonify({'error': '書類IDを指定してください'}), 400

        # セッション用の一時ディレクトリ作成
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        if not os.path.exists(base_temp_dir):
            os.makedirs(base_temp_dir, exist_ok=True)

        temp_dir = tempfile.mkdtemp(dir=base_temp_dir)
        session_id = os.path.basename(temp_dir)

        # EDINET APIからXBRLを並行ダウンロード
        api = EdinetAPI(EDINET_API_KEY)
        downloaded_files = []

        def download_single_xbrl(doc_id, retry_count=2):
            """単一XBRLをダウンロードする関数（リトライ機能付き）"""
            file_path = os.path.join(temp_dir, f"{doc_id}.zip")
            start_time = time.time()

            for attempt in range(retry_count + 1):
                success = api.download_xbrl(doc_id, file_path)

                if success:
                    elapsed = time.time() - start_time
                    return {
                        'doc_id': doc_id,
                        'file_path': file_path,
                        'status': 'success',
                        'elapsed': round(elapsed, 2),
                        'retry_count': attempt
                    }

                # リトライ前に少し待機
                if attempt < retry_count:
                    time.sleep(1)

            elapsed = time.time() - start_time
            return {
                'doc_id': doc_id,
                'status': 'failed',
                'elapsed': round(elapsed, 2),
                'retry_count': retry_count
            }

        # 並行ダウンロード実行（最大6並列 - リトライ機能でカバー）
        start_time = time.time()
        max_workers = min(6, len(doc_ids))  # 最大6並列、またはファイル数

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(download_single_xbrl, doc_id): doc_id for doc_id in doc_ids}

            for future in as_completed(futures):
                result = future.result()
                downloaded_files.append(result)

        total_elapsed = time.time() - start_time

        # 成功・失敗の統計
        success_count = sum(1 for f in downloaded_files if f['status'] == 'success')
        failed_count = len(downloaded_files) - success_count

        return jsonify({
            'session_id': session_id,
            'temp_dir': temp_dir,
            'downloaded_files': downloaded_files,
            'total_elapsed': round(total_elapsed, 2),
            'count': len(doc_ids),
            'success_count': success_count,
            'failed_count': failed_count,
            'parallel_workers': max_workers
        }), 200

    except Exception as e:
        app.logger.error(f"Error in edinet_preload: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/edinet/convert', methods=['POST'])
def edinet_download_and_convert():
    """先読みダウンロード済みのXBRLをExcelに変換"""
    import convert_xbrl_to_excel
    from flask import jsonify

    try:
        data = request.get_json()
        session_id = data.get('session_id')
        doc_ids = data.get('doc_ids', [])

        if not session_id or not doc_ids:
            return jsonify({'error': 'セッションIDと書類IDを指定してください'}), 400

        # セッションディレクトリを取得
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        temp_dir = os.path.join(base_temp_dir, session_id)

        if not os.path.exists(temp_dir):
            return jsonify({'error': 'セッションが見つかりません'}), 404

        # ダウンロード済みファイルを確認
        downloaded_paths = []
        for doc_id in doc_ids:
            file_path = os.path.join(temp_dir, f"{doc_id}.zip")
            if os.path.exists(file_path):
                downloaded_paths.append(file_path)

        if not downloaded_paths:
            return jsonify({'error': 'ダウンロード済みファイルが見つかりません'}), 404

        # Excelに変換
        out_excel = convert_xbrl_to_excel.process_xbrl_zips(downloaded_paths, output_dir=temp_dir)

        if out_excel and os.path.exists(out_excel):
            # Return JSON with the relative file path for the download API
            relative_path = os.path.relpath(out_excel, BASE_TEMP_DIR)
            return jsonify({
                "success": True,
                "file": relative_path
            })
        else:
            return jsonify({'error': 'Excelファイルの生成に失敗しました'}), 500

    except Exception as e:
        app.logger.error(f"Error in edinet_download_and_convert: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/edinet/download')
def download_excel():
    """iPad等のファイル名化け対策: ブラウザのデフォルトダウンロード機能を利用"""
    from urllib.parse import quote
    
    file_rel_path = request.args.get("file")
    if not file_rel_path:
        return "File not specified", 400

    # セキュリティチェック: ディレクトリトラバーサル防止
    if ".." in file_rel_path or file_rel_path.startswith("/") or file_rel_path.startswith("\\"):
        return "Invalid file path", 400

    file_path = os.path.normpath(os.path.join(BASE_TEMP_DIR, file_rel_path))
    
    # パスが BASE_TEMP_DIR 内にあることを確認
    if not file_path.startswith(BASE_TEMP_DIR):
        return "Access denied", 403

    if not os.path.exists(file_path):
        return "File not found", 404

    filename = os.path.basename(file_path)
    # 改行コードを除去
    filename = filename.replace('\n', '').replace('\r', '')
    encoded_filename = quote(filename)

    response = send_file(
        file_path,
        as_attachment=True,
        download_name=filename
    )

    # Content-Disposition を UTF-8 エンコードで設定（iPad等の対応）
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{encoded_filename}"

    return response


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
