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

# 前回同期状態ファイル
LAST_SYNC_FILE=".last_db_sync"

# タイムスタンプ
TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')

# ログ出力関数
log() {
    echo "[$TIMESTAMP] $1" | tee -a "$LOG_FILE"
}

# 状態ファイルから値を取得
# 引数: $1=key
get_sync_state() {
    local KEY=$1
    if [ ! -f "$LAST_SYNC_FILE" ]; then
        return 0
    fi

    grep "^${KEY}=" "$LAST_SYNC_FILE" | tail -n 1 | cut -d'=' -f2-
}

# 同期状態記録関数
# 引数: $1=mtime, $2=s211 status, $3=s217 status, $4=s63 status
write_sync_state() {
    printf 'mtime=%s\ns211=%s\ns217=%s\ns63=%s\n' "$1" "$2" "$3" "$4" > "$LAST_SYNC_FILE"
}

log "========================================"
log "EDINET DB同期開始"
log "========================================"

# DBファイル存在チェック
if [ ! -f "$DB_FILE" ]; then
    log "ERROR: $DB_FILE が見つかりません"
    exit 1
fi

# 更新チェック: 前回サーバー別の転送結果と比較
DB_MTIME=$(stat -c '%Y' "$DB_FILE")
LAST_SYNC_MTIME=""
LAST_S211_STATUS=""
LAST_S217_STATUS=""
LAST_S63_STATUS=""

if [ -f "$LAST_SYNC_FILE" ]; then
    LAST_SYNC_MTIME=$(get_sync_state "mtime")
    LAST_S211_STATUS=$(get_sync_state "s211")
    LAST_S217_STATUS=$(get_sync_state "s217")
    LAST_S63_STATUS=$(get_sync_state "s63")

    if [ "$DB_MTIME" = "$LAST_SYNC_MTIME" ]; then
        if [ "$LAST_S211_STATUS" = "success" ] && \
           [ "$LAST_S217_STATUS" = "success" ] && \
           [ "$LAST_S63_STATUS" = "success" ]; then
            log "DBファイルに変更なし、かつ全サーバーで前回成功しているため転送をスキップします。"
            log "========================================"
            exit 0
        fi

        log "DBファイルは未変更ですが、前回失敗したサーバーまたは状態不明のサーバーを再転送します。"
    fi
fi

DB_SIZE=$(ls -lh "$DB_FILE" | awk '{print $5}')
log "転送ファイル: $DB_FILE ($DB_SIZE)"

# SSH秘密鍵
SSH_KEY="$HOME/.ssh/edinet_sync_key"

# SCP転送関数
# 引数: $1=ホスト, $2=ユーザー, $3=リモートパス, $4=サーバー名
scp_transfer() {
    local HOST=$1
    local USER=$2
    local REMOTE_PATH=$3
    local SERVER_NAME=$4

    log "---"
    log "$SERVER_NAME への転送開始..."

    scp -i "$SSH_KEY" \
        -o StrictHostKeyChecking=no \
        -o ConnectTimeout=30 \
        "$DB_FILE" "${USER}@${HOST}:${REMOTE_PATH}/${DB_FILE}" 2>&1 | tee -a "$LOG_FILE"

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

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 設定ファイルが存在すれば読み込み
CONFIG_FILE="$SCRIPT_DIR/.ftp_config"
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

# s63設定
S63_HOST="${S63_HOST:-s63.xrea.com}"
S63_USER="${S63_USER:-tomo}"
S63_PASS="${S63_PASS:-YOUR_PASSWORD_HERE}"
S63_PATH="${S63_PATH:-/virtual/tomo/public_html/xbrl1.xtomo.com}"

# ==========================================
# 転送実行
# ==========================================

SUCCESS_COUNT=0
FAIL_COUNT=0
SKIP_COUNT=0
S211_RESULT="failure"
S217_RESULT="failure"
S63_RESULT="failure"

# s211への転送
if [ "$DB_MTIME" = "$LAST_SYNC_MTIME" ] && [ "$LAST_S211_STATUS" = "success" ]; then
    log "---"
    log "s211 は前回成功済みのため転送をスキップします。"
    ((SKIP_COUNT++))
    S211_RESULT="success"
else
    if scp_transfer "$S211_HOST" "$S211_USER" "$S211_PATH" "s211"; then
        ((SUCCESS_COUNT++))
        S211_RESULT="success"
    else
        ((FAIL_COUNT++))
    fi
fi

# s217への転送
if [ "$DB_MTIME" = "$LAST_SYNC_MTIME" ] && [ "$LAST_S217_STATUS" = "success" ]; then
    log "---"
    log "s217 は前回成功済みのため転送をスキップします。"
    ((SKIP_COUNT++))
    S217_RESULT="success"
else
    if scp_transfer "$S217_HOST" "$S217_USER" "$S217_PATH" "s217"; then
        ((SUCCESS_COUNT++))
        S217_RESULT="success"
    else
        ((FAIL_COUNT++))
    fi
fi

# s63への転送
if [ "$DB_MTIME" = "$LAST_SYNC_MTIME" ] && [ "$LAST_S63_STATUS" = "success" ]; then
    log "---"
    log "s63 は前回成功済みのため転送をスキップします。"
    ((SKIP_COUNT++))
    S63_RESULT="success"
else
    if scp_transfer "$S63_HOST" "$S63_USER" "$S63_PATH" "s63"; then
        ((SUCCESS_COUNT++))
        S63_RESULT="success"
    else
        ((FAIL_COUNT++))
    fi
fi

# ==========================================
# 結果サマリー
# ==========================================

log "========================================"
log "同期完了: 成功=$SUCCESS_COUNT, 失敗=$FAIL_COUNT, スキップ=$SKIP_COUNT"
log "========================================"

write_sync_state "$DB_MTIME" "$S211_RESULT" "$S217_RESULT" "$S63_RESULT"

if [ $FAIL_COUNT -gt 0 ]; then
    log "エラーが発生しました"
    exit 1
else
    log "全サーバーへの転送成功"
    exit 0
fi
