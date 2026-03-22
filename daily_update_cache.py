#!/virtual/tomo/public_html/xbrl3.xtomo.com/venv/bin/python3
# -*- coding: utf-8 -*-
"""
EDINET キャッシュ 日次更新スクリプト

毎日午前7時にcronで実行し、前日のEDINETデータを取得してedinet_cache.dbを更新
英語名辞書も自動的に更新（毎週金曜日、または辞書ファイルが無い場合）

【使い方】
  python3 daily_update_cache.py [--date YYYY-MM-DD] [--days N] [--update-english-dict] [--skip-english-dict]

【オプション】
  --date YYYY-MM-DD      : 指定日のデータを取得（デフォルト: 昨日）
  --days N               : 過去N日分を取得（デフォルト: 1）
  --update-english-dict  : 英語名辞書を強制的に更新
  --skip-english-dict    : 英語名辞書の更新をスキップ

【exit code】
  0: 成功（データ取得あり）
  1: エラー（API失敗、DB書き込み失敗）
  2: 成功（データなし）
"""

import sys
import os
import requests
from datetime import datetime, timedelta
import logging
from typing import Tuple
import csv
import json
import re
from collections import defaultdict

# スクリプトのディレクトリに移動（cronから実行されても動作するように）
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

from edinet_cache import EdinetCache

# ログ設定
LOG_FILE = "daily_update_cache.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# EDINET APIキー
from edinet_api_key import EDINET_API_KEY as API_KEY
BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"
EDINET_CODELIST_URL = "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/codelisteng/Edinetcode.zip"


def extract_katakana_reading(phonetic_name: str, japanese_name: str) -> tuple:
    """
    カナ名から主要部分を抽出

    Args:
        phonetic_name: カナ名（例: "アサヒカセイカブシキガイシャ"）
        japanese_name: 日本語名（例: "旭化成株式会社"）

    Returns:
        (カタカナ読み, 企業名主要部分) のタプル
        抽出できない場合は None
    """
    if not phonetic_name or not japanese_name:
        return None

    # カナ名から除外するパターン
    exclude_patterns = [
        r'カブシキガイシャ$',
        r'カブシキカイシャ$',
        r'ユウゲンガイシャ$',
        r'ゴウドウガイシャ$',
        r'ゴウシガイシャ$',
        r'ザイダンホウジン$',
        r'シャダンホウジン$',
        r'トクテイヒエイリカツドウホウジン$',
    ]

    kana_main = phonetic_name
    for pattern in exclude_patterns:
        kana_main = re.sub(pattern, '', kana_main)

    if len(kana_main) < 2:
        return None

    # 日本語名から企業名主要部分を抽出
    katakana_parts = re.findall(r'[ァ-ヴー]+', japanese_name)
    if katakana_parts:
        japanese_main = max(katakana_parts, key=len)
    else:
        japanese_main = re.sub(r'(株式会社|有限会社|合同会社|合資会社|財団法人|社団法人|特定非営利活動法人)$', '', japanese_name)
        japanese_main = re.sub(r'^(株式会社|有限会社|合同会社|合資会社|財団法人|社団法人|特定非営利活動法人)', '', japanese_main)

    if len(japanese_main) < 1:
        return None

    return (kana_main.lower(), japanese_main)


def extract_company_name(english_name: str, japanese_name: str) -> tuple:
    """
    英語名と日本語名から主要部分を抽出

    Args:
        english_name: 英語の企業名（例: "CANON INC."）
        japanese_name: 日本語の企業名（例: "キヤノン株式会社"）

    Returns:
        (英語主要部分(小文字), カタカナ主要部分) のタプル
        抽出できない場合は None
    """
    if not english_name or not japanese_name:
        return None

    # 英語名から企業名主要部分を抽出
    exclude_patterns = [
        r'\bCO\.,?\s*LTD\.?',
        r'\bCORPORATION',
        r'\bCORP\.?',
        r'\bINC\.?',
        r'\bLTD\.?',
        r'\bLIMITED',
        r'\bHOLDINGS?',
        r'\bGROUP',
        r'\bCOMPANY',
        r'\bK\.?K\.?',
    ]

    eng_clean = english_name.upper()
    for pattern in exclude_patterns:
        eng_clean = re.sub(pattern, '', eng_clean, flags=re.IGNORECASE)

    eng_clean = eng_clean.strip().strip(',').strip()

    # 複数単語の場合の処理
    words = eng_clean.split()
    if len(words) == 1:
        eng_main = words[0]
    elif len(words) <= 3:
        eng_main = ' '.join(words)
    else:
        eng_main = ' '.join(words[:2])

    if len(eng_main) < 2:
        return None

    # 日本語名からカタカナ部分を抽出
    katakana_parts = re.findall(r'[ァ-ヴー]+', japanese_name)
    if not katakana_parts:
        return None

    kata_main = max(katakana_parts, key=len)

    if len(kata_main) < 2:
        return None

    return (eng_main.lower(), kata_main)


def update_english_dictionary() -> Tuple[bool, int, str]:
    """
    EDINET公式英語コードリストから英語名辞書とカタカナ読み辞書を更新

    Returns:
        (success, entry_count, message)
    """
    import zipfile
    import tempfile

    logger.info("英語名辞書とカタカナ読み辞書の更新を開始...")

    try:
        # ZIPファイルをダウンロード
        logger.info(f"EDINETコードリストをダウンロード中: {EDINET_CODELIST_URL}")
        response = requests.get(EDINET_CODELIST_URL, timeout=60)

        if response.status_code != 200:
            return False, 0, f"Download failed: HTTP {response.status_code}"

        # 一時ファイルに保存して解凍
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_zip:
            tmp_zip.write(response.content)
            tmp_zip_path = tmp_zip.name

        with tempfile.TemporaryDirectory() as tmp_dir:
            # 解凍
            with zipfile.ZipFile(tmp_zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            csv_path = os.path.join(tmp_dir, 'EdinetcodeDlInfo.csv')

            if not os.path.exists(csv_path):
                return False, 0, "CSV file not found in ZIP"

            # 辞書を構築（英語名とカタカナ読みの両方）
            english_dict = {}
            english_dict_metadata = {}
            katakana_dict = {}
            katakana_dict_metadata = {}
            stats = defaultdict(int)

            with open(csv_path, 'r', encoding='cp932', errors='ignore') as f:
                reader = csv.reader(f)
                next(reader)  # メタ行
                next(reader)  # ヘッダー

                for row in reader:
                    if len(row) < 9:
                        continue

                    stats['total'] += 1
                    listing_status = row[2]
                    japanese_name = row[6]
                    english_name = row[7]
                    phonetic_name = row[8]

                    # 優先度計算（上場企業を優先）
                    is_listed = 1 if listing_status == 'Listed company' else 2

                    # 英語名辞書の構築
                    if english_name and english_name.strip() != '':
                        result = extract_company_name(english_name, japanese_name)
                        if result is not None:
                            eng_main, kata_main = result
                            priority = (is_listed, len(kata_main))

                            # 複数のキー候補を生成
                            key_candidates = [eng_main]
                            words = eng_main.split()
                            if len(words) > 1:
                                key_candidates.append(words[0])

                            for key in key_candidates:
                                if key in english_dict:
                                    existing_priority = english_dict_metadata[key]
                                    if priority < existing_priority:
                                        english_dict[key] = kata_main
                                        english_dict_metadata[key] = priority
                                        stats['eng_updated'] += 1
                                    else:
                                        stats['eng_duplicate_skipped'] += 1
                                else:
                                    english_dict[key] = kata_main
                                    english_dict_metadata[key] = priority
                                    stats['eng_added'] += 1
                        else:
                            stats['eng_extraction_failed'] += 1
                    else:
                        stats['no_english'] += 1

                    # カタカナ読み辞書の構築
                    if phonetic_name and phonetic_name.strip() != '':
                        result = extract_katakana_reading(phonetic_name, japanese_name)
                        if result is not None:
                            kana_main, japanese_main = result
                            priority = (is_listed, len(japanese_main))

                            if kana_main in katakana_dict:
                                existing_priority = katakana_dict_metadata[kana_main]
                                if priority < existing_priority:
                                    katakana_dict[kana_main] = japanese_main
                                    katakana_dict_metadata[kana_main] = priority
                                    stats['kata_updated'] += 1
                                else:
                                    stats['kata_duplicate_skipped'] += 1
                            else:
                                katakana_dict[kana_main] = japanese_main
                                katakana_dict_metadata[kana_main] = priority
                                stats['kata_added'] += 1
                        else:
                            stats['kata_extraction_failed'] += 1
                    else:
                        stats['no_phonetic'] += 1

        # 一時ZIPファイルを削除
        os.unlink(tmp_zip_path)

        # JSONファイルに保存
        english_output_file = 'english_katakana_dict.json'
        with open(english_output_file, 'w', encoding='utf-8') as f:
            json.dump(english_dict, f, ensure_ascii=False, indent=2, sort_keys=True)

        katakana_output_file = 'katakana_company_dict.json'
        with open(katakana_output_file, 'w', encoding='utf-8') as f:
            json.dump(katakana_dict, f, ensure_ascii=False, indent=2, sort_keys=True)

        logger.info(f"  総企業数: {stats['total']}社")
        logger.info(f"  【英語名辞書】")
        logger.info(f"    英語名なし: {stats['no_english']}社")
        logger.info(f"    抽出失敗: {stats['eng_extraction_failed']}社")
        logger.info(f"    辞書エントリ数: {len(english_dict)}件")
        logger.info(f"    保存先: {english_output_file}")
        logger.info(f"  【カタカナ読み辞書】")
        logger.info(f"    カナ名なし: {stats['no_phonetic']}社")
        logger.info(f"    抽出失敗: {stats['kata_extraction_failed']}社")
        logger.info(f"    辞書エントリ数: {len(katakana_dict)}件")
        logger.info(f"    保存先: {katakana_output_file}")

        return True, len(english_dict) + len(katakana_dict), "OK"

    except requests.exceptions.Timeout:
        return False, 0, "Download timeout"
    except requests.exceptions.RequestException as e:
        return False, 0, f"Network error: {e}"
    except Exception as e:
        return False, 0, f"Unexpected error: {e}"


def fetch_reports_for_date(target_date: str) -> Tuple[bool, int, str]:
    """
    指定日のEDINETデータを取得

    Args:
        target_date: YYYY-MM-DD形式の日付

    Returns:
        (success, report_count, message)
        success: 成功時True
        report_count: 取得した報告書数
        message: メッセージ
    """
    url = f"{BASE_URL}/documents.json"
    params = {"date": target_date, "type": 2}
    headers = {"Ocp-Apim-Subscription-Key": API_KEY}

    try:
        response = requests.get(url, params=params, headers=headers, timeout=30)

        if response.status_code != 200:
            return False, 0, f"HTTP {response.status_code}"

        data = response.json()

        # メタデータチェック
        metadata = data.get('metadata', {})
        if metadata.get('status') != '200':
            status = metadata.get('status')
            message = metadata.get('message', 'Unknown error')
            return False, 0, f"API Status {status}: {message}"

        # 有価証券報告書のみ抽出
        reports = []
        if data.get('results'):
            for doc in data['results']:
                # docTypeCode: 120 = 有価証券報告書
                # xbrlFlag: 1 = XBRL形式あり
                if doc.get('docTypeCode') == '120' and str(doc.get('xbrlFlag')) == '1':
                    reports.append(doc)

        return True, len(reports), "OK"

    except requests.exceptions.Timeout:
        return False, 0, "Timeout"
    except requests.exceptions.RequestException as e:
        return False, 0, f"Network error: {e}"
    except Exception as e:
        return False, 0, f"Unexpected error: {e}"


def update_cache_for_days(days: int = 1, start_date: str = None, update_english_dict: bool = None) -> int:
    """
    過去N日分のデータをキャッシュに追加

    Args:
        days: 取得日数
        start_date: 開始日（YYYY-MM-DD）。Noneの場合は昨日から
        update_english_dict: 英語辞書の更新を強制（True）、スキップ（False）、自動判定（None）

    Returns:
        exit code (0: 成功, 1: エラー, 2: データなし)
    """
    logger.info("=" * 80)
    logger.info("EDINET キャッシュ 日次更新開始")
    logger.info("=" * 80)

    # 英語名辞書の更新判定
    should_update_dict = False

    if update_english_dict is True:
        # 強制更新
        should_update_dict = True
    elif update_english_dict is False:
        # スキップ
        should_update_dict = False
    else:
        # 自動判定（週に1回、金曜日のみ、または辞書ファイルが無い場合）
        today = datetime.now()
        is_friday = today.weekday() == 4
        should_update_dict = is_friday or not os.path.exists('english_katakana_dict.json')

    if should_update_dict:
        logger.info("")
        logger.info("-" * 80)
        success, entry_count, message = update_english_dictionary()
        if success:
            logger.info(f"✓ 英語名辞書とカタカナ読み辞書を更新しました（合計{entry_count}件）")
        else:
            logger.warning(f"✗ 辞書の更新に失敗: {message}")
        logger.info("-" * 80)
        logger.info("")

    # 開始日の決定
    if start_date:
        try:
            end_date = datetime.strptime(start_date, "%Y-%m-%d")
        except ValueError:
            logger.error(f"Invalid date format: {start_date} (expected YYYY-MM-DD)")
            return 1
    else:
        # デフォルト: 昨日
        end_date = datetime.now() - timedelta(days=1)

    start_date_obj = end_date - timedelta(days=days - 1)

    logger.info(f"取得期間: {start_date_obj.strftime('%Y-%m-%d')} 〜 {end_date.strftime('%Y-%m-%d')} ({days}日分)")

    # キャッシュDB準備
    try:
        cache = EdinetCache()
    except Exception as e:
        logger.error(f"Failed to initialize cache DB: {e}")
        return 1

    # 更新前の統計
    stats_before = cache.get_cache_stats()
    logger.info(f"更新前: {stats_before['total_reports']:,}件の報告書")

    # 日ごとに取得
    total_added = 0
    total_success = 0
    total_failed = 0
    failed_dates = []

    for i in range(days):
        target_date_obj = start_date_obj + timedelta(days=i)
        target_date = target_date_obj.strftime("%Y-%m-%d")

        logger.info(f"Processing {target_date}...")

        success, report_count, message = fetch_reports_for_date(target_date)

        if success:
            if report_count > 0:
                # キャッシュに追加
                try:
                    url = f"{BASE_URL}/documents.json"
                    params = {"date": target_date, "type": 2}
                    headers = {"Ocp-Apim-Subscription-Key": API_KEY}
                    response = requests.get(url, params=params, headers=headers, timeout=30)
                    data = response.json()

                    reports = []
                    if data.get('results'):
                        for doc in data['results']:
                            if doc.get('docTypeCode') == '120' and str(doc.get('xbrlFlag')) == '1':
                                reports.append(doc)

                    cache.add_reports(reports)
                    total_added += report_count
                    total_success += 1
                    logger.info(f"  ✓ {report_count}件追加")

                except Exception as e:
                    total_failed += 1
                    failed_dates.append(target_date)
                    logger.error(f"  ✗ DB write failed: {e}")
            else:
                # データなし（正常）
                total_success += 1
                logger.info(f"  - データなし")
        else:
            total_failed += 1
            failed_dates.append(target_date)
            logger.error(f"  ✗ API error: {message}")

    # 更新後の統計
    stats_after = cache.get_cache_stats()
    logger.info("")
    logger.info("=" * 80)
    logger.info("更新完了")
    logger.info("=" * 80)
    logger.info(f"成功: {total_success}日, 失敗: {total_failed}日")
    logger.info(f"追加件数: {total_added}件")
    logger.info(f"更新後: {stats_after['total_reports']:,}件の報告書 (総企業数: {stats_after['total_companies']:,}社)")
    logger.info(f"DBサイズ: {stats_after['db_size_mb']:.2f}MB")

    if failed_dates:
        logger.warning(f"失敗した日付: {', '.join(failed_dates)}")

    # メタデータ更新
    cache.set_metadata('last_update_date', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    cache.set_metadata('last_update_added_count', str(total_added))
    cache.set_metadata('last_update_failed_count', str(total_failed))

    # Exit code決定
    if total_failed > 0:
        logger.error("エラーが発生しました（exit code: 1）")
        return 1
    elif total_added == 0:
        logger.info("新規データなし（exit code: 2）")
        return 2
    else:
        logger.info("正常終了（exit code: 0）")
        return 0


def main():
    """メイン処理"""
    import argparse

    parser = argparse.ArgumentParser(description='EDINET キャッシュ 日次更新')
    parser.add_argument('--date', type=str, help='取得開始日（YYYY-MM-DD形式、デフォルト: 昨日）')
    parser.add_argument('--days', type=int, default=1, help='取得日数（デフォルト: 1）')
    parser.add_argument('--update-english-dict', action='store_true', help='英語名辞書を強制的に更新')
    parser.add_argument('--skip-english-dict', action='store_true', help='英語名辞書の更新をスキップ')

    args = parser.parse_args()

    # 英語辞書の更新フラグを決定
    update_dict = None
    if args.update_english_dict:
        update_dict = True
    elif args.skip_english_dict:
        update_dict = False

    try:
        exit_code = update_cache_for_days(
            days=args.days,
            start_date=args.date,
            update_english_dict=update_dict
        )
        sys.exit(exit_code)
    except KeyboardInterrupt:
        logger.warning("\n中断されました")
        sys.exit(1)
    except Exception as e:
        logger.error(f"予期しないエラー: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
