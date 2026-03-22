# EDINET英語名辞書について

## 概要

企業名の英語検索を可能にするため、EDINET公式の英語コードリストから英語名→カタカナ名の辞書を自動生成しています。

## 辞書の更新方法

### 方法1: 日次更新スクリプトを使用（推奨）

日次更新スクリプト `daily_update_cache.py` には英語辞書の自動更新機能が組み込まれています。

**自動更新（毎週金曜日）**
```bash
python3 daily_update_cache.py
```
- 毎週金曜日に自動的に英語辞書を更新
- 辞書ファイルが無い場合も自動的に生成

**手動で強制更新**
```bash
python3 daily_update_cache.py --update-english-dict
```

**更新をスキップ**
```bash
python3 daily_update_cache.py --skip-english-dict
```

### 方法2: 手動更新（個別に辞書のみ更新）

```bash
python3 build_english_dict_from_edinet.py
```

これにより `english_katakana_dict.json` が最新のEDINETデータで更新されます。

## 辞書の仕様

- **ソース**: EDINET公式英語コードリスト（EdinetcodeDlInfo.csv）
- **エントリ数**: 約3,600件（2026年3月時点）
- **形式**: JSON (`english_katakana_dict.json`)
- **キー**: 英語の企業名（小文字、複数単語対応）
- **値**: カタカナの企業名（主要部分）

### 優先順位

同じ英語名で複数の企業がある場合、以下の優先順位で選択されます：

1. 上場企業を優先
2. カタカナ名が短い方を優先

### 複数単語の企業名

複数単語の企業名（例: "TOYOTA MOTOR"）は、以下の両方の形式で辞書に登録されます：

- 完全版: `"toyota motor": "トヨタ"`
- 短縮版: `"toyota": "トヨタ"`

これにより、"toyota"だけでも"toyota motor"でも検索可能です。

## 使用方法

辞書は `edinet_cache.py` の `normalize_text()` 関数で自動的に読み込まれます。

```python
from edinet_api import EdinetAPI

api = EdinetAPI('your_api_key')
results = api.search_company('canon')  # キヤノンがヒット
```

## 既知の制限事項

- 英語名が登録されていない企業（約6,800社）は検索できません
- 長音符の削除により、一部誤マッチする可能性があります（例: ricoh）
- 同じ英語名で始まる複数企業がある場合、最も優先度の高い企業のみが辞書に登録されます

## ファイル一覧

- `english_katakana_dict.json` - 生成された辞書ファイル
- `build_english_dict_from_edinet.py` - 辞書生成スクリプト
- `EdinetcodeDlInfo.csv` - EDINET公式データ（ダウンロード後）
- `Edinetcode.zip` - ダウンロードしたZIPファイル

## 更新頻度

- **自動更新**: `daily_update_cache.py` を毎日実行している場合、毎週金曜日に自動更新されます
- **手動更新**: 必要に応じて `--update-english-dict` オプションで強制更新可能
- EDINETのコードリストは随時更新されるため、週次更新で最新の企業情報を維持できます

## cronでの自動実行例

```cron
# 毎日午前7時に実行（金曜日は英語辞書も自動更新）
0 7 * * * cd /path/to/work_office_edinet_api && python3 daily_update_cache.py
```
