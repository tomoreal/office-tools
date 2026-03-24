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

    # 英語名から企業形態を除く
    english_main = english_name.upper()
    for suffix in [' INC.', ' INC', ' CORPORATION', ' CORP.', ' CORP', ' CO.,LTD.', ' CO.,LTD', ' CO.LTD.', ' LTD.', ' LTD', ' LIMITED']:
        if english_main.endswith(suffix):
            english_main = english_main[:-len(suffix)].strip()

    # カタカナ部分を抽出
    katakana_parts = re.findall(r'[ァ-ヴー]+', japanese_name)
    if not katakana_parts:
        return None

    katakana_main = max(katakana_parts, key=len)

    return (english_main.lower(), katakana_main)


def update_company_master() -> Tuple[bool, int, str]:
    """
    EDINET公式英語コードリストからEDINETの企業マスターをSQLiteに更新
    
    Returns:
        (success, entry_count, message)
    """
    import zipfile
    import tempfile
    import sqlite3

    logger.info("企業マスタ(company_master)の更新を開始...")

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

            conn = sqlite3.connect('edinet_cache.db')
            cursor = conn.cursor()

            cursor.execute("""
                CREATE TABLE IF NOT EXISTS company_master (
                    edinet_code TEXT PRIMARY KEY,
                    japanese_name TEXT,
                    english_name TEXT,
                    kana_name TEXT
                )
            """)
            
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_cm_english ON company_master(english_name)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_cm_kana ON company_master(kana_name)")

            records = []
            with open(csv_path, 'r', encoding='cp932', errors='ignore') as f:
                reader = csv.reader(f)
                next(reader)  # メタ行
                next(reader)  # ヘッダー

                for row in reader:
                    if len(row) < 9:
                        continue
                    
                    edinet_code = row[0].strip()
                    japanese_name = row[6].strip()
                    english_name = row[7].strip()
                    kana_name = row[8].strip()
                    
                    if edinet_code:
                        records.append((edinet_code, japanese_name, english_name, kana_name))
            
            cursor.executemany("""
                INSERT OR REPLACE INTO company_master 
                (edinet_code, japanese_name, english_name, kana_name) 
                VALUES (?, ?, ?, ?)
            """, records)
            
            conn.commit()
            conn.close()
            
            os.unlink(tmp_zip_path)
            
            logger.info(f"  ✓ {len(records)}件の企業マスタを更新しました")
            return True, len(records), "OK"

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
        # 自動判定（週に1回、金曜日のみ）または company_master が存在しない場合
        today = datetime.now()
        is_friday = today.weekday() == 4
        
        has_table = False
        try:
            import sqlite3
            tmp_conn = sqlite3.connect('edinet_cache.db')
            res = tmp_conn.cursor().execute("SELECT name FROM sqlite_master WHERE type='table' AND name='company_master'").fetchone()
            if res:
                has_table = True
            tmp_conn.close()
        except:
            pass

        should_update_dict = is_friday or not has_table

    if should_update_dict:
        logger.info("")
        logger.info("-" * 80)
        success, entry_count, message = update_company_master()
        if success:
            logger.info(f"✓ 企業マスタを更新しました（合計{entry_count}件）")
        else:
            logger.warning(f"✗ マスタの更新に失敗: {message}")
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

    # 10年超の古いデータを削除（EDINETは約10年分のデータしか保持していないため）
    logger.info("")
    logger.info("-" * 80)
    logger.info("10年超の古いデータを削除中...")
    try:
        deleted_count = cache.delete_old_reports(years=10)
        if deleted_count > 0:
            logger.info(f"✓ {deleted_count:,}件の古いデータを削除しました")
            # 削除後の統計
            stats_after_cleanup = cache.get_cache_stats()
            logger.info(f"削除後: {stats_after_cleanup['total_reports']:,}件の報告書")
            logger.info(f"DBサイズ: {stats_after_cleanup['db_size_mb']:.2f}MB (削減: {stats_after['db_size_mb'] - stats_after_cleanup['db_size_mb']:.2f}MB)")
        else:
            logger.info("削除対象のデータはありませんでした")
    except Exception as e:
        logger.warning(f"古いデータの削除に失敗: {e}")
    logger.info("-" * 80)

    # 過去の失敗を検出して再試行（過去7日間をチェック）
    logger.info("")
    logger.info("-" * 80)
    logger.info("データ欠落チェック（過去7日間の再試行）...")
    try:
        retry_added = 0
        retry_dates = []

        # 過去7日間の各日付をチェック
        for i in range(1, 8):
            check_date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')

            # その日付のデータがDBに存在するか確認
            import sqlite3
            conn = sqlite3.connect('edinet_cache.db')
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM securities_reports
                WHERE DATE(submit_datetime) = ?
            """, (check_date,))
            count = cursor.fetchone()[0]
            conn.close()

            # データが0件の場合、その日のデータを再取得
            if count == 0:
                success, report_count, message = fetch_reports_for_date(check_date)

                if success and report_count > 0:
                    # データ取得成功、DBに追加
                    try:
                        url = f"{BASE_URL}/documents.json"
                        params = {"date": check_date, "type": 2}
                        headers = {"Ocp-Apim-Subscription-Key": API_KEY}
                        response = requests.get(url, params=params, headers=headers, timeout=30)
                        data = response.json()

                        reports = []
                        if data.get('results'):
                            for doc in data['results']:
                                if doc.get('docTypeCode') == '120' and str(doc.get('xbrlFlag')) == '1':
                                    reports.append(doc)

                        cache.add_reports(reports)
                        retry_added += report_count
                        retry_dates.append(check_date)
                        logger.info(f"  ✓ {check_date}: {report_count}件を再取得して追加")
                    except Exception as e:
                        logger.warning(f"  ✗ {check_date}: DB書き込み失敗 - {e}")
                elif success and report_count == 0:
                    # データなし（正常）
                    pass
                else:
                    # API エラー（警告のみ、次回再試行される）
                    logger.warning(f"  ! {check_date}: API取得失敗 - {message}")

        if retry_added > 0:
            logger.info(f"✓ 過去の欠落データを{retry_added}件追加しました（{len(retry_dates)}日分）")
            logger.info(f"  対象日付: {', '.join(retry_dates)}")
        else:
            logger.info("欠落データはありませんでした")
    except Exception as e:
        logger.warning(f"欠落チェックに失敗: {e}")
    logger.info("-" * 80)

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
