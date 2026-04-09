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
import time
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

# 並行ダウンロードの排他制御用ロックファイル
DOWNLOAD_LOCK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.lock_download')


def acquire_download_lock(max_retries=2, wait_sec=3, stale_sec=60):
    """ダウンロードロックを取得する。

    複数ユーザーが同時にダウンロードすると6並列×N本になるのを防ぐため、
    ファイルロックで排他制御する。CGI環境（マルチプロセス）でも機能する。

    取得できない場合は False を即座に返す（最大 max_retries×wait_sec 秒だけ待機）。
    呼び出し側は False の場合に 503 を返し、フロントエンドがリトライする設計。

    Returns:
        True: ロック取得成功
        False: ロック取得失敗（サーバービジー）
    """
    for attempt in range(max_retries):
        # タイムスタンプが古いロックは削除（前のプロセスが異常終了した場合）
        if os.path.exists(DOWNLOAD_LOCK_FILE):
            age = time.time() - os.path.getmtime(DOWNLOAD_LOCK_FILE)
            if age > stale_sec:
                try:
                    os.remove(DOWNLOAD_LOCK_FILE)
                    app.logger.warning(f"Removed stale download lock (age={age:.0f}s)")
                except OSError:
                    pass
                continue

        # O_CREAT | O_EXCL でアトミックにファイル作成（競合回避）
        try:
            fd = os.open(DOWNLOAD_LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return True
        except FileExistsError:
            pass  # 他のプロセスが保持中

        app.logger.info(f"Download lock busy, waiting {wait_sec}s (attempt {attempt + 1}/{max_retries})")
        time.sleep(wait_sec)

    app.logger.info("Download lock busy: returning server_busy to client")
    return False


def release_download_lock():
    """ダウンロードロックを解放する。"""
    try:
        os.remove(DOWNLOAD_LOCK_FILE)
    except OSError:
        pass

# ========================================================================
# HOUSEKEEPING - ログローテーション・一時ファイル定期削除
# ========================================================================
# cronが使えない環境向け：リクエスト時にセンチネルファイルのmtimeを確認し、
# 一定間隔でまとめてハウスキーピングを実行する。

def _run_housekeeping():
    """ログローテーションと一時ファイル削除をまとめて実行する。

    センチネルファイル（temp_uploads/.last_cleanup）のmtimeで実行間隔を管理する。
    CHECK_INTERVAL_HOURS 未満なら即リターンするため、毎リクエスト呼んでもコストは低い。
    """
    CHECK_INTERVAL_HOURS = 1   # センチネルが新しければスキップ（最小チェック間隔）
    TEMP_MAX_AGE_HOURS    = 12   # この時間より古い temp_uploads サブディレクトリを削除

    sentinel = os.path.join(BASE_TEMP_DIR, '.last_cleanup')
    now = time.time()

    # チェック間隔未満ならスキップ
    if os.path.exists(sentinel):
        if now - os.path.getmtime(sentinel) < CHECK_INTERVAL_HOURS * 3600:
            return

    # センチネルを先に更新（並行リクエストの二重実行を緩和）
    try:
        open(sentinel, 'w').close()
    except OSError:
        return

    # 1. 一時ディレクトリの古いセッションを削除
    cutoff = now - TEMP_MAX_AGE_HOURS * 3600
    for name in os.listdir(BASE_TEMP_DIR):
        if name.startswith('.'):
            continue
        path = os.path.join(BASE_TEMP_DIR, name)
        if os.path.isdir(path):
            try:
                if os.path.getmtime(path) < cutoff:
                    shutil.rmtree(path)
            except Exception:
                pass

    # 2. ログローテーション（convert_xbrl_to_excel の rotate_logs_manually と二重構造）
    #    CLI直接実行時は debug_log() 内のチェックが担当するため、ここでは Web経由のみカバー。
    try:
        import convert_xbrl_to_excel
        convert_xbrl_to_excel.rotate_logs_manually(convert_xbrl_to_excel._LOG_FILE)
    except Exception:
        pass


@app.before_request
def before_request_housekeeping():
    _run_housekeeping()


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

        # 並行ダウンロード実行（最大6並列）
        # 複数ユーザーの同時リクエストで6並列×N本にならないようロックで排他制御
        start_time = time.time()
        max_workers = min(6, len(doc_ids))  # 最大6並列、またはファイル数

        if not acquire_download_lock():
            return jsonify({'error': 'server_busy', 'message': '他のユーザーがダウンロード中です。しばらくお待ちください。'}), 503

        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(download_single_xbrl, doc_id): doc_id for doc_id in doc_ids}

                for future in as_completed(futures):
                    result = future.result()
                    downloaded_files.append(result)
        finally:
            release_download_lock()

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


@app.route('/api/edinet/download-pdfs', methods=['POST'])
def edinet_download_pdfs():
    """有価証券報告書のPDFをダウンロードしてZIPで固める"""
    from edinet_api import EdinetAPI
    from flask import jsonify
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time
    import zipfile
    import re

    try:
        data = request.get_json()
        session_id = data.get('session_id')
        doc_ids = data.get('doc_ids', [])
        company_name = data.get('company_name', '')
        reports_info = data.get('reports_info', [])

        if not session_id or not doc_ids:
            return jsonify({'error': 'セッションIDと書類IDを指定してください'}), 400

        # セッションディレクトリを取得
        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        temp_dir = os.path.join(base_temp_dir, session_id)

        if not os.path.exists(temp_dir):
            return jsonify({'error': 'セッションが見つかりません'}), 404

        # PDFダウンロード用のサブディレクトリを作成
        pdf_dir = os.path.join(temp_dir, 'pdfs')
        os.makedirs(pdf_dir, exist_ok=True)

        # 報告書情報をdocIDでマッピング
        reports_map = {report['docID']: report for report in reports_info}

        # EDINET APIからPDFを並行ダウンロード
        api = EdinetAPI(EDINET_API_KEY)

        def sanitize_filename(name):
            """ファイル名として使えない文字を除去"""
            return re.sub(r'[\\/:*?"<>|]', '_', name)

        def get_period_ym(period_end):
            """期末日からYYYYMM形式を取得 (例: 2025-03-31 -> 202503)"""
            if period_end:
                return period_end.replace('-', '')[:6]
            return ''

        def download_single_pdf(doc_id, retry_count=2):
            """単一PDFをダウンロードする関数（リトライ機能付き）"""
            # 報告書情報から期末を取得
            report = reports_map.get(doc_id, {})
            period_end = report.get('periodEnd', '')
            period_ym = get_period_ym(period_end)

            # ファイル名: 有報_企業名_年月.pdf
            safe_company_name = sanitize_filename(company_name)
            if period_ym:
                pdf_filename = f"有報_{safe_company_name}_{period_ym}.pdf"
            else:
                pdf_filename = f"有報_{safe_company_name}_{doc_id}.pdf"

            file_path = os.path.join(pdf_dir, pdf_filename)
            start_time = time.time()

            for attempt in range(retry_count + 1):
                # download_type=2 でPDFを取得
                success = api.download_xbrl(doc_id, file_path, download_type=2)

                if success:
                    elapsed = time.time() - start_time
                    return {
                        'doc_id': doc_id,
                        'file_path': file_path,
                        'filename': pdf_filename,
                        'period_end': period_end,
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

        # 並行ダウンロード実行
        start_time = time.time()
        max_workers = min(6, len(doc_ids))
        downloaded_pdfs = []

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(download_single_pdf, doc_id): doc_id for doc_id in doc_ids}

            for future in as_completed(futures):
                result = future.result()
                downloaded_pdfs.append(result)

        total_elapsed = time.time() - start_time

        # 成功したPDFをZIPに固める
        success_pdfs = [pdf for pdf in downloaded_pdfs if pdf['status'] == 'success']

        if not success_pdfs:
            return jsonify({'error': 'PDFのダウンロードに失敗しました'}), 500

        # 期間を計算（最古と最新）
        period_ends = sorted([pdf['period_end'] for pdf in success_pdfs if pdf.get('period_end')])
        if period_ends:
            start_ym = get_period_ym(period_ends[0])
            end_ym = get_period_ym(period_ends[-1])
            period_range = f"{start_ym}-{end_ym}" if start_ym != end_ym else start_ym
        else:
            period_range = str(int(time.time()))

        # PDF保護解除処理
        try:
            import pikepdf
            has_pikepdf = True
        except ImportError:
            has_pikepdf = False
            app.logger.warning("pikepdf not installed - PDF protection removal skipped")

        unlocked_pdfs = []
        for pdf_info in success_pdfs:
            original_path = pdf_info['file_path']

            if has_pikepdf:
                try:
                    # 保護解除されたPDFを同じ場所に上書き保存
                    with pikepdf.open(original_path, allow_overwriting_input=True) as pdf:
                        # 保護を解除して保存（上書き）
                        pdf.save(original_path)
                    app.logger.info(f"PDF protection removed: {pdf_info['filename']}")
                    unlocked_pdfs.append(pdf_info)
                except Exception as e:
                    app.logger.warning(f"Failed to unlock PDF {pdf_info['filename']}: {e}")
                    # 解除失敗してもファイルは追加
                    unlocked_pdfs.append(pdf_info)
            else:
                # pikepdfがない場合はそのまま追加
                unlocked_pdfs.append(pdf_info)

        # ZIPファイル作成: 有報_企業名_期間.zip
        safe_company_name = sanitize_filename(company_name)
        zip_filename = f"有報_{safe_company_name}_{period_range}.zip"
        zip_path = os.path.join(temp_dir, zip_filename)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for pdf_info in unlocked_pdfs:
                file_path = pdf_info['file_path']
                arcname = pdf_info['filename']
                zipf.write(file_path, arcname=arcname)

        # 相対パスを返す
        relative_path = os.path.relpath(zip_path, BASE_TEMP_DIR)

        return jsonify({
            'success': True,
            'file': relative_path,
            'pdf_count': len(success_pdfs),
            'failed_count': len(downloaded_pdfs) - len(success_pdfs),
            'total_elapsed': round(total_elapsed, 2)
        }), 200

    except Exception as e:
        app.logger.error(f"Error in edinet_download_pdfs: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/edinet/download-xbrls', methods=['POST'])
def edinet_download_xbrls():
    """既にダウンロード済みのXBRL ZIPをまとめてZIPに固める（再ダウンロードなし）"""
    from flask import jsonify
    import zipfile
    import re

    try:
        data = request.get_json()
        session_id = data.get('session_id')
        doc_ids = data.get('doc_ids', [])
        company_name = data.get('company_name', '')
        reports_info = data.get('reports_info', [])

        if not session_id or not doc_ids:
            return jsonify({'error': 'セッションIDと書類IDを指定してください'}), 400

        base_temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_uploads')
        temp_dir = os.path.join(base_temp_dir, session_id)

        if not os.path.exists(temp_dir):
            return jsonify({'error': 'セッションが見つかりません'}), 404

        def sanitize_filename(name):
            return re.sub(r'[\\/:*?"<>|]', '_', name)

        def get_period_ym(period_end):
            if period_end:
                return period_end.replace('-', '')[:6]
            return ''

        reports_map = {report['docID']: report for report in reports_info}
        safe_company_name = sanitize_filename(company_name)

        # 既存のダウンロード済みZIPを収集
        xbrl_files = []
        for doc_id in doc_ids:
            file_path = os.path.join(temp_dir, f"{doc_id}.zip")
            if os.path.exists(file_path):
                report = reports_map.get(doc_id, {})
                period_end = report.get('periodEnd', '')
                period_ym = get_period_ym(period_end)
                if period_ym:
                    arcname = f"有報XBRL_{safe_company_name}_{period_ym}.zip"
                else:
                    arcname = f"有報XBRL_{safe_company_name}_{doc_id}.zip"
                xbrl_files.append({'file_path': file_path, 'arcname': arcname, 'period_end': period_end})

        if not xbrl_files:
            return jsonify({'error': 'ダウンロード済みのXBRLファイルが見つかりません'}), 404

        # 期間範囲を計算
        period_ends = sorted([f['period_end'] for f in xbrl_files if f.get('period_end')])
        if period_ends:
            start_ym = get_period_ym(period_ends[0])
            end_ym = get_period_ym(period_ends[-1])
            period_range = f"{start_ym}-{end_ym}" if start_ym != end_ym else start_ym
        else:
            import time
            period_range = str(int(time.time()))

        # ZIPファイル作成: 有報XBRL_企業名_期間.zip
        zip_filename = f"有報XBRL_{safe_company_name}_{period_range}.zip"
        zip_path = os.path.join(temp_dir, zip_filename)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for xbrl in xbrl_files:
                zipf.write(xbrl['file_path'], arcname=xbrl['arcname'])

        relative_path = os.path.relpath(zip_path, BASE_TEMP_DIR)

        return jsonify({
            'success': True,
            'file': relative_path,
            'xbrl_count': len(xbrl_files)
        }), 200

    except Exception as e:
        app.logger.error(f"Error in edinet_download_xbrls: {e}")
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

@app.route('/download/ppm_add_label')
def download_ppm_add_label():
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates', 'PPM_add_label.bas')
    return send_file(file_path, as_attachment=True, download_name='PPM_add_label.bas')

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
