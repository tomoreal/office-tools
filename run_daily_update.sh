#!/bin/bash
#
# EDINET キャッシュ日次更新 - cronラッパースクリプト
#
# corem15のcronから実行するためのシェルスクリプト
# daily_update_cache.py を呼び出す
#

# スクリプトのディレクトリに移動
cd /virtual/tomo/public_html/xbrl3.xtomo.com

# venv内のPython3で日次更新スクリプトを実行
/virtual/tomo/public_html/xbrl3.xtomo.com/venv/bin/python3 daily_update_cache.py >> daily_update_cache.log 2>&1

# exit codeをそのまま返す
exit $?
