# 本番環境デプロイ手順

## 前提条件
- 本番サーバー: XREA/CoreServerなどのCGI対応サーバー
- Python 3.9以上
- 仮想環境（venv）が利用可能

## デプロイ手順

### 1. ファイルのアップロード

以下のファイルを本番サーバーにアップロード：
- index.cgi
- app.py
- edinet_api.py
- convert_xbrl_to_excel.py
- csv_converter.html
- templates/ (全ファイル)
- venv/ (全体)
- requirements.txt

### 2. パーミッション設定

```bash
chmod 755 index.cgi app.py edinet_api.py convert_xbrl_to_excel.py
chmod -R 755 venv/
chmod -R 755 templates/
mkdir -p temp_uploads edinet_downloads edinet_taxonomies
chmod 777 temp_uploads edinet_downloads
```

### 3. venvパッケージ確認

必須パッケージ:
- Flask==3.0.3
- beautifulsoup4==4.12.3
- lxml==5.2.2
- pandas==2.2.2
- openpyxl==3.1.5
- requests==2.32.5

不足がある場合:
```bash
venv/bin/pip install -r requirements.txt
```

## 動作確認手順

### 1. ブラウザでアクセス
https://xbrl.xtomo.com/ (または該当URL)

### 2. EDINET API検索テスト
1. 企業名入力（例: トヨタ）
2. 検索ボタンクリック
3. 企業選択
4. 報告書選択
5. ダウンロード＆変換実行
6. Excelファイルダウンロード確認

### 3. D&D方式テスト
1. EDINETから手動でZIPダウンロード
2. 方法2にD&D
3. 変換実行
4. Excelファイルダウンロード確認

## トラブルシューティング

### 500エラー
- shebangパス確認（index.cgi 1行目）
- パーミッション確認
- パッケージインストール確認

### ModuleNotFoundError
```bash
venv/bin/pip install --target=venv/lib/python3.9/site-packages パッケージ名
```

### EDINET API動作しない
- APIキー確認（app.py 59行目）
- requests確認
- ネットワーク確認

最終更新: 2026-03-20
