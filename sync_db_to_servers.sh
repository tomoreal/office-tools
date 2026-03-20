#!/bin/bash
#
# EDINET キャッシュDB 自動同期スクリプト
#
# corem15からs211/s217へedinet_cache.dbをFTP転送
# 毎日午前7時30分にcronで実行（daily_update_cache.py実行後）
#
# 【使い方】
#   ./sync_db_to_servers.sh
#
# 【前提条件】
#   - lftp がインストールされていること（yum install lftp または apt install lftp）
#   - FTP接続情報が環境変数または設定ファイルで設定されていること
#
# 【exit code】
#   0: 全サーバーへの転送成功
#   1: 一部または全てのサーバーへの転送失敗
#

# set -e を削除（エラーがあっても続行して両方のサーバーに転送を試みる）

# スクリプトのディレクトリに移動
cd "$(dirname "$0")"

# ログファイル
LOG_FILE="sync_db_to_servers.log"

# 転送対象ファイル
DB_FILE="edinet_cache.db"

# タイムスタンプ
TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')

# ログ出力関数
log() {
    echo "[$TIMESTAMP] $1" | tee -a "$LOG_FILE"
}

log "========================================"
log "EDINET DB同期開始"
log "========================================"

# DBファイル存在チェック
if [ ! -f "$DB_FILE" ]; then
    log "ERROR: $DB_FILE が見つかりません"
    exit 1
fi

DB_SIZE=$(du -h "$DB_FILE" | cut -f1)
log "転送ファイル: $DB_FILE ($DB_SIZE)"

# FTP転送関数（lftp使用）
# 引数: $1=ホスト, $2=ユーザー, $3=パスワード, $4=リモートパス
ftp_transfer() {
    local HOST=$1
    local USER=$2
    local PASS=$3
    local REMOTE_PATH=$4
    local SERVER_NAME=$5

    log "---"
    log "$SERVER_NAME への転送開始..."

    # lftpで転送（ミラーモード、既存ファイルは上書き）
    lftp -c "
        set ftp:ssl-allow no
        set net:timeout 30
        set net:max-retries 3
        open -u $USER,$PASS $HOST
        cd $REMOTE_PATH
        put -O . $DB_FILE
        bye
    " 2>&1 | tee -a "$LOG_FILE"

    if [ ${PIPESTATUS[0]} -eq 0 ]; then
        log "✓ $SERVER_NAME への転送成功"
        return 0
    else
        log "✗ $SERVER_NAME への転送失敗"
        return 1
    fi
}

# ==========================================
# FTP接続情報設定
# ==========================================
# セキュリティのため、環境変数または別ファイルから読み込むことを推奨

# 設定ファイルが存在すれば読み込み
CONFIG_FILE=".ftp_config"
if [ -f "$CONFIG_FILE" ]; then
    source "$CONFIG_FILE"
fi

# s211設定（環境変数またはデフォルト）
S211_HOST="${S211_HOST:-s211.xrea.com}"
S211_USER="${S211_USER:-tomo}"
S211_PASS="${S211_PASS:-YOUR_PASSWORD_HERE}"
S211_PATH="${S211_PATH:-/virtual/tomo/public_html/xbrl2.xtomo.com}"

# s217設定
S217_HOST="${S217_HOST:-s217.xrea.com}"
S217_USER="${S217_USER:-tomo}"
S217_PASS="${S217_PASS:-YOUR_PASSWORD_HERE}"
S217_PATH="${S217_PATH:-/virtual/tomo/public_html/xbrl.xtomo.com}"

# ==========================================
# 転送実行
# ==========================================

SUCCESS_COUNT=0
FAIL_COUNT=0

# s211への転送
if ftp_transfer "$S211_HOST" "$S211_USER" "$S211_PASS" "$S211_PATH" "s211"; then
    ((SUCCESS_COUNT++))
else
    ((FAIL_COUNT++))
fi

# s217への転送
if ftp_transfer "$S217_HOST" "$S217_USER" "$S217_PASS" "$S217_PATH" "s217"; then
    ((SUCCESS_COUNT++))
else
    ((FAIL_COUNT++))
fi

# ==========================================
# 結果サマリー
# ==========================================

log "========================================"
log "同期完了: 成功=$SUCCESS_COUNT, 失敗=$FAIL_COUNT"
log "========================================"

if [ $FAIL_COUNT -gt 0 ]; then
    log "エラーが発生しました"
    exit 1
else
    log "全サーバーへの転送成功"
    exit 0
fi
