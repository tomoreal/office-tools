#!/virtual/tomo/public_html/xbrl3.xtomo.com/venv/bin/python3
# -*- coding: utf-8 -*-
"""
EDINET キャッシュ 日次更新スクリプト

毎日午前7時にcronで実行し、前日のEDINETデータを取得してedinet_cache.dbを更新
更新成功/失敗をログに記録し、エラー時は標準エラー出力でcronに通知

【使い方】
  python3 daily_update_cache.py [--date YYYY-MM-DD] [--days N]

【オプション】
  --date YYYY-MM-DD  : 指定日のデータを取得（デフォルト: 昨日）
  --days N           : 過去N日分を取得（デフォルト: 1）

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

# EDINET API設定
API_KEY = "6ea174edf112439da66798a6d863a95d"
BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"


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


def update_cache_for_days(days: int = 1, start_date: str = None) -> int:
    """
    過去N日分のデータをキャッシュに追加

    Args:
        days: 取得日数
        start_date: 開始日（YYYY-MM-DD）。Noneの場合は昨日から

    Returns:
        exit code (0: 成功, 1: エラー, 2: データなし)
    """
    logger.info("=" * 80)
    logger.info("EDINET キャッシュ 日次更新開始")
    logger.info("=" * 80)

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

    args = parser.parse_args()

    try:
        exit_code = update_cache_for_days(days=args.days, start_date=args.date)
        sys.exit(exit_code)
    except KeyboardInterrupt:
        logger.warning("\n中断されました")
        sys.exit(1)
    except Exception as e:
        logger.error(f"予期しないエラー: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
