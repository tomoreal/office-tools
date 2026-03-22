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
import unicodedata


def hiragana_to_katakana(text: str) -> str:
    """
    ひらがなをカタカナに変換

    Args:
        text: ひらがなテキスト

    Returns:
        カタカナテキスト
    """
    katakana = ''
    for char in text:
        # ひらがな範囲: \u3041-\u3096
        if '\u3041' <= char <= '\u3096':
            # ひらがなからカタカナへの変換（0x60足す）
            katakana += chr(ord(char) + 0x60)
        else:
            katakana += char
    return katakana


def normalize_text(text: str) -> str:
    """
    企業名検索用のテキスト正規化

    全角・半角の統一、表記ゆれの吸収を行う

    Args:
        text: 正規化するテキスト

    Returns:
        正規化されたテキスト
    """
    if not text:
        return ""

    import re

    # ひらがな→カタカナ変換（検索を統一するため）
    text = hiragana_to_katakana(text)

    # 【ステップ3】NFKC正規化（全角英数字→半角、半角カタカナ→全角など）
    normalized = unicodedata.normalize('NFKC', text)

    # 半角英数字を全角に統一
    trans_table = str.maketrans(
        'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789',
        'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９'
    )
    normalized = normalized.translate(trans_table)

    # 表記ゆれの吸収（順序が重要）
    replacements = [
        # スペースの削除（最初に実行）
        (' ', ''),
        ('　', ''),
        # 長音符の削除
        ('ー', ''),
        ('−', ''),
        ('－', ''),
        # 拗音・促音の展開
        ('フィ', 'フイ'),
        ('ティ', 'テイ'),
        ('ディ', 'デイ'),
        ('ウィ', 'ウイ'),
        # ヴの統一
        ('ヴ', 'ブ'),
        ('ゔ', 'ぶ'),
        # 小書き文字の統一（濁点処理の前に実行）
        ('ッ', 'ツ'),
        ('ャ', 'ヤ'),
        ('ュ', 'ユ'),
        ('ョ', 'ヨ'),
        # 濁点・半濁点の削除（最後に実行）
        ('ガ', 'カ'), ('ギ', 'キ'), ('グ', 'ク'), ('ゲ', 'ケ'), ('ゴ', 'コ'),
        ('ザ', 'サ'), ('ジ', 'シ'), ('ズ', 'ス'), ('ゼ', 'セ'), ('ゾ', 'ソ'),
        ('ダ', 'タ'), ('ヂ', 'チ'), ('ヅ', 'ツ'), ('デ', 'テ'), ('ド', 'ト'),
        ('バ', 'ハ'), ('ビ', 'ヒ'), ('ブ', 'フ'), ('ベ', 'ヘ'), ('ボ', 'ホ'),
        ('パ', 'ハ'), ('ピ', 'ヒ'), ('プ', 'フ'), ('ペ', 'ヘ'), ('ポ', 'ホ'),
        ('が', 'か'), ('ぎ', 'き'), ('ぐ', 'く'), ('げ', 'け'), ('ご', 'こ'),
        ('ざ', 'さ'), ('じ', 'シ'), ('ず', 'ス'), ('ぜ', 'セ'), ('ぞ', 'ソ'),
        ('だ', 'た'), ('ぢ', 'ち'), ('づ', 'つ'), ('で', 'て'), ('ど', 'と'),
        ('ば', 'は'), ('び', 'ひ'), ('ぶ', 'フ'), ('ベ', 'ヘ'), ('ボ', 'ホ'),
        ('ぱ', 'は'), ('ぴ', 'ひ'), ('ぷ', 'ふ'), ('ぺ', 'へ'), ('ぽ', 'ほ'),
    ]

    for old, new in replacements:
        normalized = normalized.replace(old, new)

    return normalized


class EdinetCache:
    """EDINET 有価証券報告書のローカルキャッシュ"""

    def __init__(self, db_path: str = None):
        """
        Args:
            db_path: SQLiteデータベースファイルのパス（Noneの場合はモジュールと同じディレクトリのedinet_cache.dbを使用）
        """
        if db_path is None:
            # モジュールの場所を基準に絶対パスを生成
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.db_path = os.path.join(base_dir, "edinet_cache.db")
        else:
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
        企業名で検索（あいまい検索対応）

        全角・半角の違い、表記ゆれを吸収して検索する

        Args:
            company_name: 検索する企業名（部分一致）

        Returns:
            企業情報リスト（EDINETコードごとにグループ化）
        """
        if not company_name:
            return []

        # 半角/全角の揺れをある程度無効化するための正規化（主に日本語での検索用）
        normalized_query = normalize_text(company_name)
        # そのままのクエリ（英語の大文字小文字や半角などをそのまま検索するため）
        raw_query = company_name.strip()
        # ひらがなをカタカナに（かな名検索用）
        kana_query = hiragana_to_katakana(raw_query)

        search_norm = f"%{normalized_query}%"
        search_raw = f"%{raw_query}%"
        search_kana = f"%{kana_query}%"

        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # company_masterがあればJOINし、english_name, kana_name, japanese_name全てを検索対象に入れる
        # また、書類種別を「有価証券報告書」のみに限定し、投資法人も除外する
        cursor.execute(f"""
            SELECT DISTINCT r.edinet_code, COALESCE(c.japanese_name, r.filer_name) as display_name, r.sec_code,
                   MAX(r.submit_datetime) as latest_submit
            FROM securities_reports r
            LEFT JOIN company_master c ON r.edinet_code = c.edinet_code
            WHERE (
               c.english_name LIKE ? COLLATE NOCASE OR
               c.kana_name LIKE ? OR
               c.japanese_name LIKE ? OR
               c.japanese_name LIKE ? OR
               r.filer_name LIKE ? OR
               r.filer_name LIKE ? OR
               r.edinet_code LIKE ? COLLATE NOCASE OR
               r.sec_code LIKE ?
            )
            AND display_name NOT LIKE '%投資法人%'
            AND r.doc_description LIKE '有価証券報告書%' 
            AND r.doc_description NOT LIKE '有価証券報告書（%'
            GROUP BY r.edinet_code
            ORDER BY latest_submit DESC
        """, (search_raw, search_kana, search_raw, search_norm, search_raw, search_norm, search_raw, search_raw))

        results = []
        for row in cursor.fetchall():
            results.append({
                'edinetCode': row['edinet_code'],
                'filerName': row['display_name'],
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
              AND doc_description LIKE '有価証券報告書%'
              AND doc_description NOT LIKE '有価証券報告書（%'
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
