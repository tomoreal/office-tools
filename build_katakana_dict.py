#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET公式コードリストからカタカナ読み辞書を構築

カナ名（例: アサヒカセイカブシキガイシャ）から企業名主要部分を抽出し、
カタカナ読み→企業名のマッピングを作成
"""

import csv
import json
import re
from collections import defaultdict


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

    # 短すぎる場合はスキップ
    if len(kana_main) < 2:
        return None

    # 日本語名から企業名主要部分を抽出
    # カタカナ部分を探す
    katakana_parts = re.findall(r'[ァ-ヴー]+', japanese_name)
    if katakana_parts:
        # カタカナが含まれている企業名（例: キヤノン株式会社）
        japanese_main = max(katakana_parts, key=len)
    else:
        # 漢字のみの企業名（例: 旭化成株式会社）
        # 「株式会社」などを除去
        japanese_main = re.sub(r'(株式会社|有限会社|合同会社|合資会社|財団法人|社団法人|特定非営利活動法人)$', '', japanese_name)
        japanese_main = re.sub(r'^(株式会社|有限会社|合同会社|合資会社|財団法人|社団法人|特定非営利活動法人)', '', japanese_main)

    if len(japanese_main) < 1:
        return None

    return (kana_main.lower(), japanese_main)


def build_katakana_dict_from_edinet(csv_path: str = 'EdinetcodeDlInfo.csv') -> dict:
    """
    EDINET CSVからカタカナ読み辞書を構築

    Returns:
        {カタカナ読み(小文字): 企業名主要部分} の辞書
    """
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
            phonetic_name = row[8]

            if not phonetic_name or phonetic_name.strip() == '':
                stats['no_phonetic'] += 1
                continue

            result = extract_katakana_reading(phonetic_name, japanese_name)
            if result is None:
                stats['extraction_failed'] += 1
                continue

            kana_main, japanese_main = result

            # 優先度計算（上場企業を優先、短い名前を優先）
            is_listed = 1 if listing_status == 'Listed company' else 2
            priority = (is_listed, len(japanese_main))

            # 重複がある場合は優先度の高い方を採用
            if kana_main in katakana_dict:
                existing_priority = katakana_dict_metadata[kana_main]
                if priority < existing_priority:
                    katakana_dict[kana_main] = japanese_main
                    katakana_dict_metadata[kana_main] = priority
                    stats['updated'] += 1
                else:
                    stats['duplicate_skipped'] += 1
            else:
                katakana_dict[kana_main] = japanese_main
                katakana_dict_metadata[kana_main] = priority
                stats['added'] += 1

    return katakana_dict, stats


def main():
    """メイン処理"""
    print("=== EDINET公式カタカナ読み辞書の構築 ===\n")

    # 辞書構築
    katakana_dict, stats = build_katakana_dict_from_edinet()

    # 統計表示
    print(f"総企業数: {stats['total']}社")
    print(f"カナ名なし: {stats['no_phonetic']}社")
    print(f"抽出失敗: {stats['extraction_failed']}社")
    print(f"辞書追加: {stats['added']}件")
    print(f"辞書更新: {stats['updated']}件")
    print(f"重複スキップ: {stats['duplicate_skipped']}件")
    print(f"\n最終辞書エントリ数: {len(katakana_dict)}件\n")

    # サンプル表示
    print("=== サンプル（最初の30件） ===")
    for i, (kana, name) in enumerate(sorted(katakana_dict.items())[:30], 1):
        print(f"{i:3d}. {kana:30s} → {name}")

    if len(katakana_dict) > 30:
        print(f"... 他 {len(katakana_dict) - 30}件")

    # JSONファイルに保存
    output_file = 'katakana_company_dict.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(katakana_dict, f, ensure_ascii=False, indent=2, sort_keys=True)

    print(f"\n辞書を {output_file} に保存しました")
    print("\nこのファイルは edinet_cache.py で読み込まれます。")


if __name__ == "__main__":
    main()
