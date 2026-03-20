#!/virtual/tomo/public_html/xbrl2.xtomo.com/venv/bin/python3.9
# s217.xrea.com 用 #!/virtual/tomo/public_html/xbrl.xtomo.com/venv/bin/python3.9
# makoto用: #!/virtual/tomo/public_html/makoto.xtomo.com/xbrl2excel/venv/bin/python3

# -*- coding: utf-8 -*-

"""
XBRL to Excel Converter - CGI Entry Point

【プログラム構成】
このファイルは以下の機能ブロックで構成されています:

1. ENVIRONMENT SETUP (1-23行)
   - シバン設定（サーバー環境別）
   - マルチスレッド問題回避
   - パス設定（アプリケーションディレクトリ、仮想環境）
   - 将来の分割先: deploy/cgi/index.cgi

2. IMPORT PHASE (24-31行)
   - Flask アプリケーションのインポート
   - エラーハンドリング（診断情報出力）
   - 将来の分割先: deploy/cgi/index.cgi

3. EXECUTION PHASE (33-41行)
   - CGIハンドラーによるアプリケーション実行
   - エラーハンドリング（診断情報出力）
   - 将来の分割先: deploy/cgi/index.cgi

【設計思想】
- app.py の薄いラッパー（CGI環境での実行専用）
- LiteSpeed/Apache CGI環境での動作を保証
- 仮想環境（venv）のパス設定を動的に処理
- エラー発生時は診断情報を出力してデバッグを容易化

【依存関係】
- app.py (Flask Application)
- venv/lib/python3.x/site-packages (Virtual Environment)

【注意事項】
- サーバー環境に応じて1行目または2行目のシバンを調整すること
- OPENBLAS_NUM_THREADS=1 でマルチスレッド問題を回避（LiteSpeed対策）
"""

# 注意: サーバー環境（s217等）に合わせて、1行目または2行目のシバンを調整してください。

import sys
import os
import traceback

# ========================================================================
# ENVIRONMENT SETUP
# ========================================================================
# 【将来の分割先】deploy/cgi/index.cgi

# LiteSpeedサーバー（コアサーバー等）でのマルチスレッド問題を回避
os.environ['OPENBLAS_NUM_THREADS'] = "1"

# アプリケーションのディレクトリをパスに追加
app_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, app_dir)

# 仮想環境 (venv) のパッケージパスを追加 (3.10と3.9の両方に対応)
for py_ver in ['3.10', '3.9']:
    venv_site_packages = os.path.join(app_dir, f'venv/lib/python{py_ver}/site-packages')
    if os.path.exists(venv_site_packages):
        sys.path.insert(0, venv_site_packages)

# ========================================================================
# IMPORT PHASE
# ========================================================================
# 【将来の分割先】deploy/cgi/index.cgi

try:
    from app import app
    from wsgiref.handlers import CGIHandler
except Exception:
    print("Content-Type: text/plain; charset=utf-8\n")
    print("--- Diagnostic Info: Error during CGI Initialization (Import Phase) ---")
    print(traceback.format_exc())
    sys.exit(0)

# ========================================================================
# EXECUTION PHASE
# ========================================================================
# 【将来の分割先】deploy/cgi/index.cgi

# CGIとして実行
if __name__ == '__main__':
    try:
        CGIHandler().run(app)
    except Exception:
        print("Content-Type: text/plain; charset=utf-8\n")
        print("--- Diagnostic Info: Error during CGI Execution ---")
        print(traceback.format_exc())
