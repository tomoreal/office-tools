# EDINET タクソノミ辞書の自動更新

このプロジェクトでは、EDINET公式タクソノミから勘定科目辞書を自動生成・更新できます。

## 概要

- **タクソノミソース**: [EDINET公式 勘定科目リスト](https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/download/ESE140115.xlsx)
- **生成ファイル**: `edinet_taxonomy_dict.py`
- **項目数**: 約1,959項目（EDINET公式1,907 + カスタム52）

## ファイル構成

```
.
├── update_edinet_taxonomy.py      # 自動更新スクリプト
├── edinet_taxonomy_dict.py        # 生成された辞書ファイル
├── edinet_taxonomy_elements.xlsx  # ダウンロードされたタクソノミ
├── .edinet_taxonomy.hash          # ファイル変更検知用ハッシュ
└── convert_xbrl_to_excel.py       # メインスクリプト（辞書を使用）
```

## 使い方

### 手動更新

タクソノミ辞書を手動で更新する場合：

```bash
python3 update_edinet_taxonomy.py
```

強制的に更新する場合（ファイルが変更されていなくても）：

```bash
python3 update_edinet_taxonomy.py --force
```

### 自動更新の設定

#### 方法1: cronジョブ（Linux/Mac）

毎月1日の午前3時に自動更新する例：

```bash
# crontabを編集
crontab -e

# 以下を追加
0 3 1 * * cd /home/tomo/work_office_new_phase && python3 update_edinet_taxonomy.py >> /tmp/edinet_update.log 2>&1
```

#### 方法2: systemd timer（Linux）

1. サービスファイルを作成: `/etc/systemd/system/edinet-taxonomy-update.service`

```ini
[Unit]
Description=EDINET Taxonomy Dictionary Update
After=network.target

[Service]
Type=oneshot
User=tomo
WorkingDirectory=/home/tomo/work_office_new_phase
ExecStart=/usr/bin/python3 update_edinet_taxonomy.py
StandardOutput=journal
StandardError=journal
```

2. タイマーファイルを作成: `/etc/systemd/system/edinet-taxonomy-update.timer`

```ini
[Unit]
Description=EDINET Taxonomy Dictionary Update Timer
Requires=edinet-taxonomy-update.service

[Timer]
OnCalendar=monthly
Persistent=true

[Install]
WantedBy=timers.target
```

3. 有効化

```bash
sudo systemctl daemon-reload
sudo systemctl enable edinet-taxonomy-update.timer
sudo systemctl start edinet-taxonomy-update.timer

# 状態確認
sudo systemctl status edinet-taxonomy-update.timer
```

#### 方法3: Windows タスクスケジューラ

1. タスクスケジューラを開く
2. 「基本タスクの作成」を選択
3. トリガー: 「毎月」を選択し、1日を指定
4. 操作: 「プログラムの開始」
5. プログラム: `python3`
6. 引数: `update_edinet_taxonomy.py`
7. 開始: `C:\path\to\work_office_new_phase`

## 更新の仕組み

1. **ダウンロード**: EDINET公式サイトから最新タクソノミExcelファイルをダウンロード
2. **変更検知**: ファイルのSHA256ハッシュで変更を検知
3. **辞書生成**: Excelから1,907項目を抽出し、52個のカスタムマッピングを追加
4. **ファイル出力**: `edinet_taxonomy_dict.py`を生成
5. **ハッシュ保存**: 次回の変更検知用にハッシュを保存

## カスタムマッピング

以下の項目はEDINET公式タクソノミに含まれないため、カスタムマッピングとして追加されています：

- IFRS要素のバリエーション（例: `RevenueIFRS`, `ProfitLossIFRS`）
- 短縮形（例: `Profit`, `NetIncome`）
- 財務諸表名（例: `ConsolidatedBalanceSheet`）

カスタムマッピングを追加・変更する場合は、`update_edinet_taxonomy.py`の`custom_mappings`辞書を編集してください。

## トラブルシューティング

### ダウンロード失敗

```
✗ Download failed: HTTP Error 404: Not Found
```

→ EDINETのURL構造が変更された可能性があります。最新のURLを確認して`EDINET_TAXONOMY_URL`を更新してください。

### インポートエラー

```
ModuleNotFoundError: No module named 'openpyxl'
```

→ openpyxlをインストール:
```bash
pip install openpyxl
```

### 権限エラー

```
PermissionError: [Errno 13] Permission denied
```

→ ファイルの書き込み権限を確認してください：
```bash
chmod +w edinet_taxonomy_dict.py
```

## 更新履歴の確認

生成されたファイルには更新日時が記録されます：

```bash
head -20 edinet_taxonomy_dict.py
```

出力例：
```python
"""
EDINET Taxonomy Dictionary

Auto-generated from EDINET Official Taxonomy:
https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/download/ESE140115.xlsx

Generated: 2026-03-18 23:45:00
Total: 1,959 items
- EDINET Official Taxonomy: 1,907 items
- Custom Mappings (IFRS variants, etc.): 52 items
"""
```

## 更新頻度の推奨

- **通常**: 月に1回（EDINETタクソノミは通常年1-2回更新）
- **新年度開始時**: 4月と12月は必ず実行（新タクソノミリリース時期）
- **緊急時**: 重要な会計基準変更時

## サポート

問題が発生した場合は、以下を確認してください：

1. ログファイルの確認（cronジョブの場合）
2. `update_edinet_taxonomy.py --force`で強制更新を試行
3. EDINETの公式サイトでタクソノミの最新情報を確認

## ライセンス

このスクリプトはEDINET公式タクソノミを使用しています。タクソノミの著作権は金融庁に帰属します。
