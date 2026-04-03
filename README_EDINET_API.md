# EDINET API連携機能 - 完成版ドキュメント

## 概要
EDINET APIを使用して、企業名検索から有価証券報告書のXBRLダウンロード、Excel変換まで一貫して自動化するシステムです。

## 実装済み機能

### 1. EDINET API検索機能
- **企業名検索**: 企業名（部分一致）で過去180日分の提出書類から企業を検索
- **有価証券報告書一覧取得**: 選択企業の過去5年分の有価証券報告書を取得
- **XBRL自動ダウンロード**: 選択した報告書のXBRLファイルを自動ダウンロード
- **Excel自動変換**: ダウンロードしたXBRLを自動的にExcelに変換
- **XBRL ZIP一括ダウンロード**: ダウンロードしたXBRLファイルをZIP形式で一括ダウンロード

### 2. 高度なExcel出力機能
- **セグメント情報 (PPM分析)**: 
  - 直近期と5年前のデータを比較可能なバブルチャート（PPM分析用）を自動生成。
  - 日本基準およびIFRSの両方に対応。報告セグメント合計や調整項目の除外等の高度なデータ整備。
- **人的資本・ダイバーシティ情報**:
  - 従業員数や女性管理職比率などの指標に加えて、会社名列を自動抽出。
  - XBRLタグ付けされていない連結子会社名等のプレーンテキスト（HTML）からのフォールバック取得。
- **財務諸表の並び順調整**:
  - 財務諸表（損益計算書等）の項目の並び順を実際の報告書に準じて正確に再現。

### 3. ユーザーインターフェース
- **Webフロントエンド (Flask)**
  - 企業名入力→検索→企業選択→報告書選択→ダウンロード＆変換まですべてブラウザ上で完結
  - 手動ダウンロードしたZIPファイルをD&Dでアップロードする従来の方法もサポート
  - XBRLファイルのZIP一括ダウンロード機能
  - CSV変換ツールや各種変換用ブックマークレット等のシステム提供

### 4. データキャッシュ・自動更新・バッチ処理機能
- **ローカルキャッシュによる高速化**: API応答や企業情報をSQLiteやローカルファイルでキャッシュ (`edinet_cache.py`, `build_cache.py`)
- **日次データ自動更新処理**: クーロン（Cron）によるバッチ実行を利用し、日々のEDINET情報をローカルに取得・差分更新 (`daily_update_cache.py`, `run_daily_update.sh`)
- **企業マスタおよびタクソノミ情報の自動更新**: EDINET側のマスタやタクソノミの変更を検知しシステムへ自動反映 (`update_company_master.py`, `update_edinet_taxonomy.py`)
- **サーバー間DB同期**: 更新されたキャッシュDBを各運用サーバーへ自動同期・展開 (`sync_db_to_servers.sh`)

### 5. ユーティリティ・支援ツール
- **Excel関連VBScript / マクロ**: 複数Excelファイルの結合（`エクセル結合.vbs`）や、財務データの横展開処理（`財務データ横展開ツール.vbs`）、さらにPPM分析用グラフにラベルを自動付与するマクロ (`PPM_add_label.bas`) 等、実務のExcel操作を補助するツール群

## ファイル構成

```
work_office_edinet_api/
├── アプリケーション本体・コアロジック
│   ├── app.py                           # Flaskアプリ（UI提供、基本機能）
│   ├── convert_xbrl_to_excel.py         # XBRL→Excel変換の主要コントローラー
│   ├── segment_analysis.py              # セグメント情報解析（PPMチャート生成等）
│   ├── financial_analysis.py            # 財務諸表解析（並び順調整等）
│   ├── diversity_analysis.py            # 人的資本・ダイバーシティ情報取得
│   └── edinet_api.py                    # EDINET API関連のリクエスト処理
├── キャッシュ・データ更新・バッチ処理
│   ├── edinet_cache.py                  # API応答や企業情報などのキャッシュ管理
│   ├── build_cache.py                   # キャッシュのビルド処理
│   ├── daily_update_cache.py            # 日次でのキャッシュ更新処理
│   ├── update_company_master.py         # 企業マスタの更新処理
│   ├── update_edinet_taxonomy.py        # タクソノミ情報の自動更新
│   ├── build_english_dict.py            # 英語辞書の構築ユーティリティ
│   ├── run_daily_update.sh              # バッチ実行用シェルスクリプト
│   └── sync_db_to_servers.sh            # サーバー間でのDB同期スクリプト
├── 設定・定義ファイル群
│   ├── edinet_api_key.py                # APIキーの管理
│   ├── edinet_taxonomy_dict.py          # タクソノミ辞書定義（勘定科目マッピング用）
│   ├── edinet_taxonomy_dict_clean.py    # 整理済みタクソノミ辞書
│   ├── requirements.txt                 # 依存パッケージリスト
│   ├── .edinet_api_key_config           # APIキー実設定ファイル（非公開・Git管理外）
│   ├── .ftp_config                      # FTP同期用設定ファイル（非公開・Git管理外）
│   ├── .ftp_config.template             # FTP設定のテンプレート
│   ├── .htaccess                        # Webサーバー設定ファイル
│   └── .edinet_taxonomy.hash            # タクソノミ更新検知用ハッシュ
├── Webフロントエンド・UI
│   ├── index.cgi                        # CGIエントリポイント
│   ├── csv_converter.html               # CSV変換ツール
│   ├── style.css                        # 共通スタイルシート
│   ├── app2.js                          # フロントエンドのスクリプト
│   ├── favicon.ico                      # ファビコン
│   ├── static/
│   │   └── js/xrea_ad_handler.js        # XREA広告のハンドリング用スクリプト
│   └── templates/
│       ├── index.html                   # XBRL変換UI（API検索・ダウンロードボタン等）
│       ├── bookmarklets.html            # ブックマークレット配布用ページ
│       ├── csv_bookmarklets.html        # CSV変換用ブックマークレット
│       └── PPM_add_label.bas            # PPMグラフにラベルを追加するVBAマクロ
├── ドキュメント群
│   ├── README.md                        # 全体ドキュメント
│   ├── README_EDINET_API.md             # 本ドキュメント
│   ├── README_EDINET_TAXONOMY.md        # タクソノミ関連の解説
│   ├── CLAUDE.md                        # AIアシスタント用コンテキスト・ルール
│   ├── CRON_SETUP.md                    # クーロン設定手順
│   ├── DEPLOY.md                        # デプロイ手順
│   ├── TEST_CHECKLIST.md                # テスト項目チェックリスト
│   ├── 将来の分割案.md                  # コード分割のリファクタリング案
│   ├── PPM_prior_lookup_plan.md         # PPM分析の事前検索計画
│   └── PPMグラフ修正.md                 # PPMグラフ修正についてのメモ
└── ユーティリティ・その他
    ├── エクセル結合.vbs                  # 複数Excelファイルを結合するVBScript
    ├── 財務データ横展開ツール.vbs        # ダウンロードしたデータを展開するVBScript
    └── convert_xbrl_to_excel.py.backup_before_common_dict_update # 変更前のバックアップ
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
```

### 2. EDINET APIキーの設定
プロジェクトのルートディレクトリに `.edinet_api_key_config` というファイルを作成し、取得したEDINET APIキーを1行だけ記載してください。または、環境変数 `EDINET_API_KEY` で指定することも可能です。
```bash
echo "あなたのAPIキー" > .edinet_api_key_config
```

### 3. 初期キャッシュの構築
検索および各種処理を高速に行うため、初回起動時には関連情報を取得・構築する必要があります。
```bash
# 企業マスタ・過去の提出書類情報の構築
venv/bin/python3 build_cache.py

# 取引所タクソノミ情報の構築
venv/bin/python3 update_edinet_taxonomy.py
```

### 4. 日次自動更新バッチの設定（本番環境用）
日々提出される有価証券報告書の情報をシステムに反映させるため、クーロン（Cron）で自動更新スクリプトを登録します。
```bash
# 例: 毎日早朝に自動実行
# crontab -e の末尾などに以下を追記 (パスは実環境に合わせる)
0 6 * * * /home/tomo/work_office_edinet_api/run_daily_update.sh
```

### 5. サーバーの起動

**ローカル開発環境**:
```bash
venv/bin/python3 app.py
# http://localhost:8000 にアクセス
```

**本番環境（CGIでの運用）**:
```bash
chmod 755 index.cgi
# Webサーバー経由で index.cgi へアクセスすることでFlaskアプリが起動
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
  - 連結財務諸表（IFRS/日本基準）: 実際の報告書に準じた正しい表示順で出力
  - 単体財務諸表
  - セグメント情報 & PPM分析（直近期・5年前比較バブルチャート）
  - 人的資本・ダイバーシティ情報（従業員数・会社名等の自動判別）
  - 各種注記

### システム運用
- **一時ファイル削除機能 (ハウスキーピング)**: 変換プロセス中に生成される古いZIPファイル等の一時ファイルを自動的にクリーンアップ

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
- 最終更新日: 2026-04-04

## ライセンス
プロジェクトのライセンスに準拠
