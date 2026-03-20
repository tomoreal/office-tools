#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET 有価証券報告書キャッシュモジュール

企業名検索を高速化するため、過去10年分の有価証券報告書をローカルDBにキャッシュする
日付ベースのEDINET APIの制約を回避する
"""

import sqlite3
import json
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import os


class EdinetCache:
    """EDINET 有価証券報告書のローカルキャッシュ"""

    def __init__(self, db_path: str = "edinet_cache.db"):
        """
        Args:
            db_path: SQLiteデータベースファイルのパス
        """
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        """データベース初期化"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # 有価証券報告書テーブル
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS securities_reports (
                doc_id TEXT PRIMARY KEY,
                edinet_code TEXT NOT NULL,
                filer_name TEXT NOT NULL,
                sec_code TEXT,
                doc_description TEXT,
                submit_datetime TEXT,
                period_start TEXT,
                period_end TEXT,
                doc_type_code TEXT,
                xbrl_flag TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(doc_id)
            )
        """)

        # 企業名検索用インデックス
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_filer_name
            ON securities_reports(filer_name)
        """)

        # EDINETコード検索用インデックス
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_edinet_code
            ON securities_reports(edinet_code)
        """)

        # 提出日時検索用インデックス
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_submit_datetime
            ON securities_reports(submit_datetime DESC)
        """)

        # キャッシュメタデータテーブル
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS cache_metadata (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        conn.commit()
        conn.close()

    def add_reports(self, reports: List[Dict]):
        """
        有価証券報告書をキャッシュに追加

        Args:
            reports: 報告書リスト（EDINET API レスポンスから抽出したもの）
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        for report in reports:
            cursor.execute("""
                INSERT OR REPLACE INTO securities_reports
                (doc_id, edinet_code, filer_name, sec_code, doc_description,
                 submit_datetime, period_start, period_end, doc_type_code, xbrl_flag)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report.get('docID'),
                report.get('edinetCode'),
                report.get('filerName'),
                report.get('secCode', ''),
                report.get('docDescription', ''),
                report.get('submitDateTime', ''),
                report.get('periodStart', ''),
                report.get('periodEnd', ''),
                report.get('docTypeCode', '120'),
                report.get('xbrlFlag', '1')
            ))

        conn.commit()
        conn.close()

    def search_by_company_name(self, company_name: str) -> List[Dict]:
        """
        企業名で検索

        Args:
            company_name: 検索する企業名（部分一致）

        Returns:
            企業情報リスト（EDINETコードごとにグループ化）
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT DISTINCT edinet_code, filer_name, sec_code,
                   MAX(submit_datetime) as latest_submit
            FROM securities_reports
            WHERE filer_name LIKE ?
            GROUP BY edinet_code
            ORDER BY latest_submit DESC
        """, (f'%{company_name}%',))

        results = []
        for row in cursor.fetchall():
            results.append({
                'edinetCode': row['edinet_code'],
                'filerName': row['filer_name'],
                'secCode': row['sec_code'] or '',
                'latest_submit': row['latest_submit']
            })

        conn.close()
        return results

    def get_reports_by_edinet_code(self, edinet_code: str, limit: int = 10) -> List[Dict]:
        """
        EDINETコードで有価証券報告書を取得

        Args:
            edinet_code: EDINETコード
            limit: 取得件数（デフォルト10件）

        Returns:
            有価証券報告書リスト（新しい順）
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT doc_id, doc_description, submit_datetime,
                   period_start, period_end
            FROM securities_reports
            WHERE edinet_code = ?
            ORDER BY submit_datetime DESC
            LIMIT ?
        """, (edinet_code, limit))

        results = []
        for row in cursor.fetchall():
            results.append({
                'docID': row['doc_id'],
                'docDescription': row['doc_description'],
                'submitDateTime': row['submit_datetime'],
                'periodStart': row['period_start'],
                'periodEnd': row['period_end']
            })

        conn.close()
        return results

    def get_cache_stats(self) -> Dict:
        """キャッシュ統計情報を取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM securities_reports")
        total_reports = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(DISTINCT edinet_code) FROM securities_reports")
        total_companies = cursor.fetchone()[0]

        cursor.execute("SELECT MIN(submit_datetime), MAX(submit_datetime) FROM securities_reports")
        date_range = cursor.fetchone()

        conn.close()

        return {
            'total_reports': total_reports,
            'total_companies': total_companies,
            'oldest_report': date_range[0],
            'newest_report': date_range[1],
            'db_size_mb': os.path.getsize(self.db_path) / 1024 / 1024 if os.path.exists(self.db_path) else 0
        }

    def set_metadata(self, key: str, value: str):
        """メタデータ設定"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT OR REPLACE INTO cache_metadata (key, value, updated_at)
            VALUES (?, ?, CURRENT_TIMESTAMP)
        """, (key, value))

        conn.commit()
        conn.close()

    def get_metadata(self, key: str) -> Optional[str]:
        """メタデータ取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute("SELECT value FROM cache_metadata WHERE key = ?", (key,))
        result = cursor.fetchone()

        conn.close()
        return result[0] if result else None
