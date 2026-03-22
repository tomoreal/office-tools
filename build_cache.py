#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET キャッシュ構築スクリプト

過去10年分の有価証券報告書をEDINET APIから取得してローカルDBに保存
"""

import requests
from datetime import datetime, timedelta
import time
from edinet_cache import EdinetCache

# EDINET APIキー
from edinet_api_key import EDINET_API_KEY as API_KEY
BASE_URL = "https://api.edinet-fsa.go.jp/api/v2"

def build_cache(years: int = 10, sampling_days: int = 7):
    """
    キャッシュ構築

    Args:
        years: 過去何年分取得するか
        sampling_days: サンプリング間隔（日数）
    """
    cache = EdinetCache()

    # 開始日と終了日
    end_date = datetime.now() - timedelta(days=1)  # 昨日まで
    start_date = end_date - timedelta(days=365 * years)

    days_diff = (end_date - start_date).days
    total_requests = days_diff // sampling_days + 1

    print(f"=== EDINET キャッシュ構築開始 ===")
    print(f"期間: {start_date.strftime('%Y-%m-%d')} 〜 {end_date.strftime('%Y-%m-%d')}")
    print(f"サンプリング間隔: {sampling_days}日")
    print(f"予想リクエスト数: 約{total_requests}回")
    print(f"予想所要時間: 約{total_requests * 0.5 / 60:.1f}分")
    print()

    processed = 0
    added_reports = 0
    errors = 0

    for i in range(0, days_diff + 1, sampling_days):
        target_date = (start_date + timedelta(days=i)).strftime("%Y-%m-%d")

        url = f"{BASE_URL}/documents.json"
        params = {"date": target_date, "type": 2}
        headers = {"Ocp-Apim-Subscription-Key": API_KEY}

        try:
            response = requests.get(url, params=params, headers=headers, timeout=30)

            if response.status_code == 200:
                data = response.json()

                # メタデータチェック
                if data.get('metadata', {}).get('status') != '200':
                    processed += 1
                    continue

                # 有価証券報告書のみ抽出
                reports = []
                if data.get('results'):
                    for doc in data['results']:
                        if doc.get('docTypeCode') == '120' and str(doc.get('xbrlFlag')) == '1':
                            reports.append(doc)

                if reports:
                    cache.add_reports(reports)
                    added_reports += len(reports)

                processed += 1

                # 進捗表示
                if processed % 50 == 0:
                    progress = processed / total_requests * 100
                    print(f"進捗: {processed}/{total_requests} ({progress:.1f}%) - 累計{added_reports}件追加")

            else:
                errors += 1
                print(f"エラー: {target_date} - HTTP {response.status_code}")

            # レート制限対策
            time.sleep(0.2)

        except Exception as e:
            errors += 1
            print(f"例外: {target_date} - {e}")
            time.sleep(1)

    # 完了メッセージ
    print()
    print("=== キャッシュ構築完了 ===")
    print(f"処理日数: {processed}日分")
    print(f"追加件数: {added_reports}件")
    print(f"エラー数: {errors}件")

    # 統計情報
    stats = cache.get_cache_stats()
    print()
    print("=== キャッシュ統計 ===")
    print(f"総報告書数: {stats['total_reports']}件")
    print(f"総企業数: {stats['total_companies']}社")
    print(f"期間: {stats['oldest_report']} 〜 {stats['newest_report']}")
    print(f"DBサイズ: {stats['db_size_mb']:.2f}MB")

    # メタデータ更新
    cache.set_metadata('last_build_date', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    cache.set_metadata('build_start_date', start_date.strftime('%Y-%m-%d'))
    cache.set_metadata('build_end_date', end_date.strftime('%Y-%m-%d'))


if __name__ == "__main__":
    # 過去10年分を全日検索（サンプリングなし、約3650回のリクエスト）
    # 所要時間: 約30-40時間
    build_cache(years=10, sampling_days=1)
