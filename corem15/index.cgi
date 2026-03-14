#!/virtual/tomo/public_html/makoto.xtomo.com/xbrl2excel/venv/bin/python3
# for s217  #!/virtual/tomo/public_html/xbrl.xtomo.com/venv/bin/python3.9

# -*- coding: utf-8 -*-
# 注意: サーバー環境（s217等）に合わせて、1行目または2行目のシバンを調整してください。

import sys
import os
import traceback

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

try:
    from app import app
    from wsgiref.handlers import CGIHandler
except Exception:
    print("Content-Type: text/plain; charset=utf-8\n")
    print("--- Diagnostic Info: Error during CGI Initialization (Import Phase) ---")
    print(traceback.format_exc())
    sys.exit(0)

# CGIとして実行
if __name__ == '__main__':
    try:
        CGIHandler().run(app)
    except Exception:
        print("Content-Type: text/plain; charset=utf-8\n")
        print("--- Diagnostic Info: Error during CGI Execution ---")
        print(traceback.format_exc())
