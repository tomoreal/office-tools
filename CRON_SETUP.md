# EDINET 自動更新システム - crom15設定手順

## 概要

corem15サーバーで毎日午前7時にEDINETキャッシュDBを自動更新し、s211/s217に配布します。

## システム構成

```
[毎日 7:00] daily_update_cache.py
    ↓ 前日のEDINETデータ取得
    ↓ edinet_cache.db を更新

[毎日 7:30] sync_db_to_servers.sh
    ↓ 更新されたDBをFTP転送
    ↓
    ├─→ s211 (xbrl2.xtomo.com)
    └─→ s217 (xbrl.xtomo.com)
```

## 前提条件

### 1. 必須パッケージのインストール

corem15にSSHでログインして実行:

```bash
# lftpインストール（FTP転送用）
# CentOS/RHEL系の場合
sudo yum install lftp

# Debian/Ubuntu系の場合
# sudo apt install lftp
```

### 2. Python仮想環境の確認

```bash
cd /virtual/tomo/public_html/xbrl3.xtomo.com
source venv/bin/activate
pip list | grep requests  # requestsがインストールされているか確認
```

---

## セットアップ手順

### 1. スクリプトのアップロード

以下のファイルをcorem15の `/virtual/tomo/public_html/xbrl3.xtomo.com/` にアップロード:

```bash
# ローカルから実行
scp daily_update_cache.py run_daily_update.sh sync_db_to_servers.sh \
    corem15:/virtual/tomo/public_html/xbrl3.xtomo.com/
```

### 2. 実行権限の付与

```bash
ssh corem15
cd /virtual/tomo/public_html/xbrl3.xtomo.com

chmod +x daily_update_cache.py
chmod +x run_daily_update.sh
chmod +x sync_db_to_servers.sh
```

### 3. FTP接続情報の設定

セキュリティのため、FTP接続情報を環境変数ファイルに保存します:

```bash
cd /virtual/tomo/public_html/xbrl3.xtomo.com
nano .ftp_config
```

以下の内容を記述（パスワードは実際のものに変更）:

```bash
# s211 FTP設定
S211_HOST="s211.xrea.com"
S211_USER="tomo"
S211_PASS="YOUR_S211_PASSWORD"
S211_PATH="/virtual/tomo/public_html/xbrl2.xtomo.com"

# s217 FTP設定
S217_HOST="s217.xrea.com"
S217_USER="tomo"
S217_PASS="YOUR_S217_PASSWORD"
S217_PATH="/virtual/tomo/public_html/xbrl.xtomo.com"
```

保存後、パーミッション設定（重要！）:

```bash
chmod 600 .ftp_config  # 自分だけ読み書き可能
```

### 4. 動作テスト

#### テスト1: 日次更新スクリプト

```bash
cd /virtual/tomo/public_html/xbrl3.xtomo.com
./daily_update_cache.py --days 1
```

期待される出力:
```
================================================================================
EDINET キャッシュ 日次更新開始
================================================================================
取得期間: 2026-03-20 〜 2026-03-20 (1日分)
更新前: 72,798件の報告書
Processing 2026-03-20...
  ✓ 15件追加
================================================================================
更新完了
================================================================================
成功: 1日, 失敗: 0日
追加件数: 15件
更新後: 72,813件の報告書 (総企業数: 4,523社)
DBサイズ: 26.54MB
正常終了（exit code: 0）
```

#### テスト2: FTP転送スクリプト

```bash
./sync_db_to_servers.sh
```

期待される出力:
```
[2026-03-21 07:30:15] ========================================
[2026-03-21 07:30:15] EDINET DB同期開始
[2026-03-21 07:30:15] ========================================
[2026-03-21 07:30:15] 転送ファイル: edinet_cache.db (26M)
[2026-03-21 07:30:15] ---
[2026-03-21 07:30:15] s211 への転送開始...
[2026-03-21 07:30:25] ✓ s211 への転送成功
[2026-03-21 07:30:25] ---
[2026-03-21 07:30:25] s217 への転送開始...
[2026-03-21 07:30:35] ✓ s217 への転送成功
[2026-03-21 07:30:35] ========================================
[2026-03-21 07:30:35] 同期完了: 成功=2, 失敗=0
[2026-03-21 07:30:35] ========================================
[2026-03-21 07:30:35] 全サーバーへの転送成功
```

---

## cron設定

### 1. crontabの編集

```bash
ssh corem15
crontab -e
```

### 2. cron設定を追加

以下の2行を追加:

```cron
# EDINET キャッシュ自動更新（毎日午前7時）
0 7 * * * /virtual/tomo/public_html/xbrl3.xtomo.com/run_daily_update.sh

# EDINET キャッシュDB自動配布（毎日午前7時30分）
30 7 * * * /virtual/tomo/public_html/xbrl3.xtomo.com/sync_db_to_servers.sh
```

保存して終了（vi: `:wq`, nano: Ctrl+O → Enter → Ctrl+X）

### 3. cron設定の確認

```bash
crontab -l
```

---

## 監視とメンテナンス

### 1. ログファイルの確認

```bash
cd /virtual/tomo/public_html/xbrl3.xtomo.com

# 更新ログ
tail -50 daily_update_cache.log

# 転送ログ
tail -50 sync_db_to_servers.log
```

### 2. 定期的な確認項目

**毎週月曜日に確認推奨:**

```bash
# 1. 過去7日分のログチェック
tail -200 daily_update_cache.log | grep "更新完了"

# 2. DB更新日時の確認
ls -lh edinet_cache.db

# 3. DBサイズの確認（異常な増加がないか）
du -h edinet_cache.db
```

### 3. エラー発生時の対処

#### エラー1: 「API error」が連続3日以上
```bash
# EDINET APIキーの有効性確認
# API_KEYが期限切れの可能性 → app.py, daily_update_cache.py のAPI_KEY更新
```

#### エラー2: 「FTP転送失敗」
```bash
# FTP接続情報の確認
cat .ftp_config

# 手動でFTPテスト
lftp -u tomo,PASSWORD s211.xrea.com
```

#### エラー3: 「DB write failed」
```bash
# DBファイルの権限確認
ls -l edinet_cache.db
chmod 644 edinet_cache.db

# DB破損チェック
sqlite3 edinet_cache.db "PRAGMA integrity_check;"
```

---

## トラブルシューティング

### Q1. cronが実行されない

**確認項目:**
```bash
# cronサービスが動いているか
systemctl status cron  # または crond

# cronログ確認（CentOS/RHEL系）
tail -50 /var/log/cron

# スクリプトの絶対パス確認
which python3
/virtual/tomo/public_html/xbrl3.xtomo.com/venv/bin/python3 --version
```

### Q2. EDINET APIからデータが取得できない

**確認項目:**
```bash
# 手動実行してエラー内容確認
cd /virtual/tomo/public_html/xbrl3.xtomo.com
./daily_update_cache.py --days 1

# APIキーの有効性確認（ブラウザで実行）
# https://api.edinet-fsa.go.jp/api/v2/documents.json?date=2026-03-20&type=2
# Header: Ocp-Apim-Subscription-Key: YOUR_API_KEY
```

### Q3. FTP転送が遅い

**対処法:**
```bash
# 圧縮転送に変更（sync_db_to_servers.sh を編集）
# lftpの代わりにscpを使う（より高速）

# または、転送頻度を減らす（毎日→週1回など）
```

---

## 参考: cron時刻の変更

毎日午前7時から別の時刻に変更したい場合:

```cron
# 午前5時に変更する場合
0 5 * * * cd /virtual/tomo/public_html/xbrl3.xtomo.com && ...

# 午前9時30分に変更する場合
30 9 * * * cd /virtual/tomo/public_html/xbrl3.xtomo.com && ...

# 毎週月曜日のみ実行する場合
0 7 * * 1 cd /virtual/tomo/public_html/xbrl3.xtomo.com && ...
```

cron時刻フォーマット: `分 時 日 月 曜日`

---

最終更新: 2026-03-21
