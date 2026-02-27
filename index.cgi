#!/virtual/tomo/public_html/makoto.xtomo.com/xbrl2excel/venv/bin/python3
# -*- coding: utf-8 -*-
# 注意: 上記のパスは実際のユーザー名と設置パスに合わせて書き換えてください。

import sys
import os

# LiteSpeedサーバー（コアサーバー等）でのマルチスレッド問題を回避
os.environ['OPENBLAS_NUM_THREADS'] = "1"

# アプリケーションのディレクトリをパスに追加
app_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, app_dir)

# 仮想環境 (venv) のパッケージパスを追加
venv_site_packages = os.path.join(app_dir, 'venv/lib/python3.10/site-packages')
if os.path.exists(venv_site_packages):
    sys.path.insert(0, venv_site_packages)

from wsgiref.handlers import CGIHandler
from app import app

# CGIとして実行
if __name__ == '__main__':
    CGIHandler().run(app)
