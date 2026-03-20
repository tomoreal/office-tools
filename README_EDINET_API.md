# EDINET API連携機能 - 完成版ドキュメント

## 概要
EDINET APIを使用して、企業名検索から有価証券報告書のXBRLダウンロード、Excel変換まで一貫して自動化するシステムです。

## 実装済み機能

### 1. EDINET API検索機能
- **企業名検索**: 企業名（部分一致）で過去180日分の提出書類から企業を検索
- **有価証券報告書一覧取得**: 選択企業の過去5年分の有価証券報告書を取得
- **XBRL自動ダウンロード**: 選択した報告書のXBRLファイルを自動ダウンロード
- **Excel自動変換**: ダウンロードしたXBRLを自動的にExcelに変換

### 2. ユーザーインターフェース
- **方法1: EDINET API検索** (推奨)
  - 企業名入力→検索→企業選択→報告書選択→ダウンロード＆変換
  - すべてブラウザ上で完結

- **方法2: 手動アップロード** (従来の方法を維持)
  - EDINETから手動ダウンロードしたZIPファイルをD&Dでアップロード

## ファイル構成

```
work_office_edinet_api/
├── edinet_api.py                    # EDINET APIクライアント (新規)
├── test_edinet_api.py               # APIテストスクリプト (新規)
├── app.py                           # Flaskアプリ (EDINET API統合済み)
├── test_server.py                   # テスト用サーバー (新規)
├── convert_xbrl_to_excel.py         # XBRL→Excel変換ロジック
├── index.cgi                        # CGIエントリポイント
├── csv_converter.html               # CSV変換ツール (旧index.html)
├── templates/
│   └── index.html                   # XBRL変換UI (EDINET検索機能追加済み)
├── edinet_downloads/                # API経由でダウンロードしたXBRL保存先
├── temp_uploads/                    # 一時ファイル保存先
├── requirements.txt                 # 依存パッケージリスト
└── EDINET_API仕様調査結果.md       # API仕様ドキュメント
```

## API エンドポイント

### 1. 企業検索
**エンドポイント**: `POST /api/edinet/search`

**リクエスト**:
```json
{
  "company_name": "トヨタ自動車"
}
```

**レスポンス**:
```json
{
  "results": [
    {
      "edinetCode": "E02144",
      "filerName": "トヨタ自動車株式会社",
      "secCode": "72030",
      "latest_submit": "2026-03-13 10:08"
    }
  ]
}
```

### 2. 有価証券報告書一覧取得
**エンドポイント**: `POST /api/edinet/reports`

**リクエスト**:
```json
{
  "edinet_code": "E02144",
  "start_date": "2021-01-01",  // オプション
  "end_date": "2026-12-31"     // オプション
}
```

**レスポンス**:
```json
{
  "results": [
    {
      "docID": "S100LO6W",
      "docDescription": "有価証券報告書－第117期...",
      "submitDateTime": "2021-06-24 15:00",
      "periodStart": "2020-04-01",
      "periodEnd": "2021-03-31"
    }
  ]
}
```

### 3. XBRL自動ダウンロード＆Excel変換
**エンドポイント**: `POST /api/edinet/convert`

**リクエスト**:
```json
{
  "doc_ids": ["S100LO6W", "S100XYZ1"]  // 複数選択可
}
```

**レスポンス**: Excelファイル（バイナリ）

**ファイル名例**: `XBRL_横展開_トヨタ自動車.xlsx`

## セットアップ方法

### 1. 依存パッケージのインストール
```bash
# venv環境の作成
python3 -m venv venv

# パッケージのインストール
venv/bin/pip install -r requirements.txt

# または手動インストール
venv/bin/pip install Flask beautifulsoup4 lxml pandas openpyxl requests
```

### 2. EDINET APIキーの設定
環境変数または[app.py](app.py:59)で直接設定：
```python
EDINET_API_KEY = os.environ.get('EDINET_API_KEY', 'あなたのAPIキー')
```

### 3. サーバーの起動

**ローカル開発環境**:
```bash
venv/bin/python3 app.py
# http://localhost:8000 にアクセス
```

**本番環境（CGI）**:
```bash
# index.cgiを通してFlaskアプリが起動
```

## 動作確認済みテスト

### ✅ API単体テスト
```bash
# 企業検索
curl -X POST http://localhost:8000/api/edinet/search \
  -H "Content-Type: application/json" \
  -d '{"company_name":"トヨタ"}'

# 有価証券報告書取得
curl -X POST http://localhost:8000/api/edinet/reports \
  -H "Content-Type: application/json" \
  -d '{"edinet_code":"E02144"}'
```

### ✅ Excel変換テスト
```bash
venv/bin/python3 convert_xbrl_to_excel.py edinet_downloads/S100LO6W.zip
# → XBRL_横展開_トヨタ自動車.xlsx (42KB) 生成成功
```

## 技術仕様

### EDINET API
- **バージョン**: v2
- **ベースURL**: `https://api.edinet-fsa.go.jp/api/v2`
- **認証**: APIキー（Subscription-Key）
- **レート制限**: 要確認

### XBRL処理
- **対応形式**: Inline XBRL (iXBRL)
- **パーサー**: BeautifulSoup4 + lxml
- **タクソノミ**: 自動取得・キャッシュ

### Excel出力
- **ライブラリ**: openpyxl
- **シート構成**:
  - 連結財務諸表（IFRS/日本基準）
  - 単体財務諸表
  - セグメント情報
  - 各種注記

## トラブルシューティング

### bs4がインストールされていない
```bash
venv/bin/pip install --target=venv/lib/python3.12/site-packages beautifulsoup4 lxml
```

### XBRLパースエラー
- タクソノミキャッシュを削除: `rm -rf edinet_taxonomies/`
- 再度実行するとタクソノミが再ダウンロードされる

### 企業が検索できない
- 検索期間を延長: デフォルト180日 → 365日以上
- 企業名の表記を変更: 「トヨタ」「トヨタ自動車」など

## 今後の改善案

1. **期間指定UI**: フロントエンドに開始日・終了日の入力フィールドを追加
2. **プログレス表示**: 長時間処理の進捗をリアルタイム表示
3. **エラーハンドリング**: より詳細なエラーメッセージとリトライ機能
4. **キャッシュ機能**: 検索結果と報告書一覧のキャッシュ
5. **バッチ処理**: 複数企業の一括ダウンロード＆変換

## 作成者
- Makoto Tomo
- 作成日: 2026-03-20

## ライセンス
プロジェクトのライセンスに準拠
