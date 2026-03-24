# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

EDINETのXBRL財務データをExcelに変換するWebアプリケーション。EDINET APIで有価証券報告書を検索・ダウンロードし、XBRLをExcel（横展開形式）に変換する。XREA/CoreServer上のCGI環境で動作。

## Running Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run Flask dev server
python app.py
# → http://localhost:8000
```

## Key Scripts

```bash
# キャッシュDB初期構築（初回のみ・過去10年分取得）
python build_cache.py

# キャッシュDB日次更新（前日分追加）
python daily_update_cache.py

# 指定日範囲の更新
python daily_update_cache.py --date 2026-03-20 --days 7

# 英語名辞書を強制更新
python daily_update_cache.py --update-english-dict
```

## Architecture

### Request Flow
```
ブラウザ → index.cgi (CGI entry) → app.py (Flask) → convert_xbrl_to_excel.py (Core)
```

**2通りの変換フロー:**
1. **D&D方式**: ZIPアップロード → `POST /` → `process_xbrl_zips()`
2. **EDINET API方式**: 企業検索 → 報告書一覧 → 先読みDL → 変換
   - `POST /api/edinet/search` → `POST /api/edinet/reports` → `POST /api/edinet/preload` → `POST /api/edinet/convert` → `GET /api/edinet/download`

### Core Files

| ファイル | 役割 |
|---|---|
| `app.py` | Flaskルーティング（薄いラッパー） |
| `index.cgi` | CGIエントリポイント（シバンはサーバー別に要調整） |
| `convert_xbrl_to_excel.py` | XBRLパース・Excel生成エンジン（約3700行） |
| `edinet_api.py` | EDINET APIクライアント |
| `edinet_cache.py` | SQLiteキャッシュ管理（企業検索の高速化） |
| `edinet_cache.db` | キャッシュDB（企業名・報告書一覧を格納） |
| `edinet_taxonomy_dict.py` | XBRLタクソノミ→日本語名称変換辞書 |
| `daily_update_cache.py` | 日次キャッシュ更新スクリプト（cron用） |
| `build_cache.py` | キャッシュ初期構築スクリプト |

### `convert_xbrl_to_excel.py` の層構造

1. **INFRASTRUCTURE LAYER** (〜380行): ログ・ファイル操作
2. **TAXONOMY LAYER** (〜905行): タクソノミ管理・解析
3. **XBRL LAYER** (〜1412行): XBRLパース・コンテキスト解析
4. **CORE LAYER** (〜3650行): メインパイプライン `process_xbrl_zips()`
5. **CLI LAYER** (3652〜): コマンドライン引数

### キャッシュシステム

- `edinet_cache.db`: SQLite（過去10年分の有価証券報告書）
- 企業名検索はDBから高速検索（EDINET APIに毎回問い合わせしない）
- EDINET APIは書類一覧取得（日付ベース）のためキャッシュが必要
- **自動更新フロー**: corem15でcron実行 → DB更新 → FTPでs211/s217に配布

### 設定ファイル

- `edinet_api_key.py`: APIキー（`.gitignore`対象）
- `.ftp_config`: FTP接続情報（パーミッション600必須）
- `edinet_api_key.py`が存在しない環境では`edinet_api_key.py.template`から作成

## Deployment

本番環境はXREA/CoreServerのCGI。詳細は `DEPLOY.md`、cron設定は `CRON_SETUP.md` 参照。

`index.cgi` の1行目シバンをサーバー環境に合わせて変更する:
```python
#!/virtual/tomo/public_html/xbrl2.xtomo.com/venv/bin/python3.9  # s211用
#!/virtual/tomo/public_html/xbrl.xtomo.com/venv/bin/python3.9   # s217用
```

LiteSpeedのマルチスレッド問題回避のため `OPENBLAS_NUM_THREADS=1` を設定済み（`index.cgi`/`app.py` 冒頭）。

## Branch Strategy

- `new-phase`: メインブランチ
- `edinet_api`: EDINET API連携機能の開発ブランチ（現在のブランチ）
