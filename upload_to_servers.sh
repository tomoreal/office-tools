#!/bin/bash
# EDINET API あいまい検索機能と英語辞書更新機能をサーバーにアップロード

set -e  # エラーが発生したら停止

echo "=== EDINET API ファイルアップロード ==="
echo ""

# アップロード先のベースディレクトリ
S211_DIR="/virtual/tomo/public_html/xbrl2.xtomo.com"
S217_DIR="/virtual/tomo/public_html/xbrl.xtomo.com"
COREM15_DIR="/virtual/tomo/public_html/xbrl3.xtomo.com"

FILES=(
#    "README_english_dict.md"
#    "edinet_cache.py"
#    "edinet_cache.db"
#    "daily_update_cache.py"
#    "build_english_dict_from_edinet.py"
#    "english_katakana_dict.json"
#    "build_katakana_dict.py"
#    "katakana_company_dict.json"
#    ".edinet_api_key_config"
#    "edinet_api_key.py"
#    "app.py"
#    "build_cache.py"
#    "daily_update_cache.py"
#    "convert_xbrl_to_excel.py"
#    "templates"
#    "app.py"
#    "convert_xbrl_to_excel.py"
#    "edinet_api.py"
#    "edinet_cache.py"
    "convert_xbrl_to_excel.py"

)

# s211へアップロード
echo "--- s211 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp -r "${FILES[@]}" "s211:${S211_DIR}/"
echo "✓ s211 完了"
echo ""

# s217へアップロード
echo "--- s217 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp -r "${FILES[@]}" "s217:${S217_DIR}/"
echo "✓ s217 完了"
echo ""

# corem15へアップロード
echo "--- corem15 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp -r "${FILES[@]}" "corem15:${COREM15_DIR}/"
echo "✓ corem15 完了"
echo ""

echo "=== すべてのアップロード完了 ==="
