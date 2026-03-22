#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英語名→カタカナ名の辞書を自動構築するスクリプト

キャッシュDBから企業名を抽出し、英語名とカタカナ名のマッピングを作成
"""

import sqlite3
import re
import json

def extract_english_katakana_pairs(db_path: str = 'edinet_cache.db'):
    """
    キャッシュDBから英語名とカタカナ名のペアを抽出

    Returns:
        dict: {英語名(小文字): カタカナ名} の辞書
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # すべての企業名を取得
    cursor.execute("""
        SELECT DISTINCT filer_name
        FROM securities_reports
        ORDER BY filer_name
    """)

    english_dict = {}

    for row in cursor.fetchall():
        filer_name = row[0]

        # 英語部分を抽出（大文字+小文字の連続、最低2文字）
        # 例: "Sony Corporation" → "Sony", "TOYOTA MOTOR" → "TOYOTA"
        english_parts = re.findall(r'\b[A-Za-z]{2,}(?:\s+[A-Za-z]+)*\b', filer_name)

        # カタカナ部分を抽出（最も長いカタカナ連続）
        katakana_parts = re.findall(r'[ァ-ヴー]+', filer_name)

        if english_parts and katakana_parts:
            # 最も長い英語部分とカタカナ部分を使用
            english = max(english_parts, key=len).strip().lower()
            katakana = max(katakana_parts, key=len)

            # 一般的な単語を除外
            exclude_words = {'corporation', 'company', 'limited', 'holdings', 'inc', 'co', 'ltd'}
            if english in exclude_words:
                continue

            # 既存のマッピングがあれば、より短いカタカナを優先
            if english in english_dict:
                if len(katakana) < len(english_dict[english]):
                    english_dict[english] = katakana
            else:
                english_dict[english] = katakana

    conn.close()
    return english_dict

def main():
    """メイン処理"""
    print("=== 英語名→カタカナ名辞書の自動生成 ===\n")

    # 辞書抽出
    english_dict = extract_english_katakana_pairs()

    # 結果表示
    print(f"抽出された辞書エントリ数: {len(english_dict)}件\n")

    # 主要企業のみ表示
    print("=== 主要企業（アルファベット順） ===")
    for english, katakana in sorted(english_dict.items())[:30]:
        print(f"  '{english}': '{katakana}',")

    if len(english_dict) > 30:
        print(f"  ... 他 {len(english_dict) - 30}件")

    # JSONファイルに保存
    output_file = 'english_katakana_dict.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(english_dict, f, ensure_ascii=False, indent=2)

    print(f"\n辞書を {output_file} に保存しました")
    print(f"\nedinet_cache.py の normalize_text() 関数内の")
    print(f"english_to_katakana 辞書をこのファイルで置き換えることができます。")

if __name__ == "__main__":
    main()
