#!/bin/bash
# EDINET API あいまい検索機能と英語辞書更新機能をサーバーにアップロード

set -e  # エラーが発生したら停止

echo "=== EDINET API ファイルアップロード ==="
echo ""

# アップロード先のベースディレクトリ
S211_DIR="/virtual/tomo/public_html/xbrl2.xtomo.com"
S217_DIR="/virtual/tomo/public_html/xbrl.xtomo.com"
COREM15_DIR="/virtual/tomo/public_html/xbrl3.xtomo.com"

# アップロードするファイル
FILES=(
#    "edinet_cache.py"
#    "daily_update_cache.py"
#    "build_english_dict_from_edinet.py"
#    "README_english_dict.md"
#    "build_katakana_dict.py"
#    ".edinet_api_key_config"
#    "edinet_api_key.py"
#    "english_katakana_dict.json"
#    "katakana_company_dict.json"
#    "app.py"
#    "build_cache.py"
#    "daily_update_cache.py"
    "convert_xbrl_to_excel.py"
)

# s211へアップロード
echo "--- s211 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp "${FILES[@]}" "s211:${S211_DIR}/"
echo "✓ s211 完了"
echo ""

# s217へアップロード
echo "--- s217 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp "${FILES[@]}" "s217:${S217_DIR}/"
echo "✓ s217 完了"
echo ""

# corem15へアップロード
echo "--- corem15 へアップロード ---"
echo "  Uploading: ${FILES[*]}"
scp "${FILES[@]}" "corem15:${COREM15_DIR}/"
echo "✓ corem15 完了"
echo ""

echo "=== すべてのアップロード完了 ==="
echo ""
echo "次のステップ："
echo "1. 各サーバーで英語辞書が正しく読み込まれるか確認"
echo "2. cronジョブが正常に動作するか確認（daily_update_cache.py）"
echo ""
