#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET API モジュール

EDINET APIへのアクセスを提供する本番用モジュール
企業検索、有価証券報告書検索、XBRLダウンロード機能を提供

【将来の分割先】api/edinet/client.py
"""

import requests
import os
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from edinet_cache import EdinetCache


class EdinetAPI:
    """EDINET API クライアント（キャッシュベース）"""

    def __init__(self, api_key: str, cache_db_path: str = None):
        """
        Args:
            api_key: EDINET APIキー
            cache_db_path: キャッシュDBのパス
        """
        self.api_key = api_key
        self.base_url = "https://api.edinet-fsa.go.jp/api/v2"
        self.cache = EdinetCache(cache_db_path)

    def search_company(self, company_name: str, search_days: int = 180) -> List[Dict]:
        """
        企業名で企業を検索（キャッシュDBから高速検索）

        Args:
            company_name: 検索する企業名（部分一致）
            search_days: 未使用（互換性のため残存）

        Returns:
            企業情報のリスト
        """
        # キャッシュDBから検索
        return self.cache.search_by_company_name(company_name)

    def get_securities_reports(
        self,
        edinet_code: str,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        limit: int = 10
    ) -> List[Dict]:
        """
        特定企業の有価証券報告書を取得（キャッシュDBから高速取得）

        Args:
            edinet_code: EDINETコード（例: 'E02144'）
            start_date: 未使用（互換性のため残存）
            end_date: 未使用（互換性のため残存）
            limit: 取得件数（デフォルト10件）

        Returns:
            有価証券報告書のリスト
            [
                {
                    'docID': 'S100LO6W',
                    'docDescription': '有価証券報告書－第117期...',
                    'submitDateTime': '2021-06-24 15:00',
                    'periodStart': '2020-04-01',
                    'periodEnd': '2021-03-31'
                },
                ...
            ]
        """
        # キャッシュDBから取得
        return self.cache.get_reports_by_edinet_code(edinet_code, limit)

    def download_xbrl(
        self,
        doc_id: str,
        output_path: str,
        download_type: int = 1
    ) -> bool:
        """
        XBRL ZIPファイルをダウンロード

        Args:
            doc_id: 書類ID（例: 'S100LO6W'）
            output_path: 保存先ファイルパス
            download_type: ダウンロード種別
                1: 提出本文書及び監査報告書（XBRL含む）← 推奨
                2: PDF
                3: 代替書面・添付文書
                4: 英文ファイル
                5: CSV形式のXBRL

        Returns:
            成功時True、失敗時False
        """
        url = f"{self.base_url}/documents/{doc_id}"
        params = {
            "type": download_type
        }
        headers = {
            "Ocp-Apim-Subscription-Key": self.api_key
        }

        try:
            response = requests.get(url, params=params, headers=headers, timeout=60)

            if response.status_code == 200:
                # 出力ディレクトリを作成
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                # ZIPファイルとして保存
                with open(output_path, 'wb') as f:
                    f.write(response.content)

                return True
            else:
                return False

        except Exception:
            return False


# 便利な関数（後方互換性のため）
def search_company(api_key: str, company_name: str, search_days: int = 180) -> List[Dict]:
    """企業検索（モジュールレベル関数）"""
    api = EdinetAPI(api_key)
    return api.search_company(company_name, search_days)


def get_securities_reports(
    api_key: str,
    edinet_code: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None
) -> List[Dict]:
    """有価証券報告書検索（モジュールレベル関数）"""
    api = EdinetAPI(api_key)
    return api.get_securities_reports(edinet_code, start_date, end_date)


def download_xbrl(
    api_key: str,
    doc_id: str,
    output_path: str,
    download_type: int = 1
) -> bool:
    """XBRLダウンロード（モジュールレベル関数）"""
    api = EdinetAPI(api_key)
    return api.download_xbrl(doc_id, output_path, download_type)
