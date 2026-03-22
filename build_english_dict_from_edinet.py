#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET公式英語コードリストから英語名辞書を構築

https://disclosure2dl.edinet-fsa.go.jp/searchdocument/codelisteng/Edinetcode.zip
から英語名→日本語名のマッピングを抽出
"""

import os
import csv
import json
import re
from collections import defaultdict


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
    # 除外する一般的な語句
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

    # 1単語: そのまま使用
    if len(words) == 1:
        eng_main = words[0]
    # 2-3単語: 全体を使用（スペース区切り）
    elif len(words) <= 3:
        eng_main = ' '.join(words)
    else:
        # 4単語以上: 最初の2単語を使用
        eng_main = ' '.join(words[:2])

    # 短すぎる場合はスキップ
    if len(eng_main) < 2:
        return None

    # 会社法人等格を表す言葉を削除し、日本語の企業名全体を採用する
    # カタカナのみを抽出すると、漢字主体の企業（本田技研工業など）が漏れるため
    jp_exclude = [
        '株式会社', '有限会社', '合同会社', '合名会社', '合資会社', 
        '相互会社', '特定目的会社', '投資法人', '一般社団法人', '公益財団法人',
        '（株）', '（有）'
    ]
    kata_main = japanese_name
    for term in jp_exclude:
        kata_main = kata_main.replace(term, '')
    kata_main = kata_main.strip()

    # 短すぎる場合はスキップ
    if len(kata_main) < 2:
        return None

    return (eng_main.lower(), kata_main)


def build_english_dict_from_edinet(csv_path: str = 'EdinetcodeDlInfo.csv') -> dict:
    """
    EDINET CSVから英語名辞書を構築

    Returns:
        {英語名(小文字): カタカナ名} の辞書
    """
    english_dict = {}
    english_dict_metadata = {}  # 優先度管理用
    all_candidates = defaultdict(list)  # 共通プレフィックス探索用
    stats = defaultdict(int)

    with open(csv_path, 'r', encoding='cp932', errors='ignore') as f:
        reader = csv.reader(f)

        # メタ行とヘッダーをスキップ
        next(reader)  # メタ行
        next(reader)  # ヘッダー

        for row in reader:
            if len(row) < 9:
                continue

            edinet_code = row[0]
            submitter_type = row[1]  # 提出者種別
            listing_status = row[2]  # 上場区分
            japanese_name = row[6]   # Submitter Name
            english_name = row[7]    # Submitter Name (alphabetic)

            stats['total'] += 1

            # 英語名が存在しない場合はスキップ
            if not english_name or english_name.strip() == '':
                stats['no_english'] += 1
                continue

            # 企業名を抽出
            result = extract_company_name(english_name, japanese_name)
            if result is None:
                stats['extraction_failed'] += 1
                continue

            eng_main, kata_main = result

            # 優先度の計算（上場企業を優先、本来の英語名が短いものを優先、短いカタカナ名を優先）
            # priority値が小さいほど優先度が高い
            is_listed = 1 if listing_status == 'Listed company' else 2
            priority = (is_listed, len(eng_main.split()), len(kata_main))

            # 複数のキー候補を生成
            # 1. 完全な企業名（例: "toyota motor"）
            # 2. 最初の単語のみ（例: "toyota"）
            key_candidates = [eng_main]
            words = eng_main.split()
            if len(words) > 1:
                key_candidates.append(words[0])

            for key in key_candidates:
                all_candidates[key].append(kata_main)
                # 重複がある場合は優先度の高い方を採用
                if key in english_dict:
                    existing_priority = english_dict_metadata[key]
                    if priority < existing_priority:
                        english_dict[key] = kata_main
                        english_dict_metadata[key] = priority
                        stats['updated'] += 1
                    else:
                        stats['duplicate_skipped'] += 1
                else:
                    english_dict[key] = kata_main
                    english_dict_metadata[key] = priority
                    stats['added'] += 1

        # 共通接頭辞を使用した汎用キーワードの上書き（例: nissan -> 日産、toyota -> トヨタ を抽出）
        for key, list_names in all_candidates.items():
            if len(list_names) > 1:
                unique_names = list(set(list_names))
                if len(unique_names) > 1:
                    # Toyota問題（「トヨタ」と「豊田」が混在すると共通接頭辞がなくなる）に対応
                    # 先頭2文字でグループ化し、最大のグループを優先する
                    prefix_groups = defaultdict(list)
                    for name in unique_names:
                        if len(name) >= 2:
                            prefix_groups[name[:2]].append(name)

                    if prefix_groups:
                        # 要素数が一番多いグループを取得
                        best_group = max(prefix_groups.values(), key=len)
                        
                        # 2社以上からなるグループの場合は、その内部で共通プレフィックスを計算
                        if len(best_group) > 1:
                            prefix = os.path.commonprefix(best_group)
                            if len(prefix) >= 2:
                                english_dict[key] = prefix
                                stats['prefix_overrides'] += 1

    return english_dict, stats


def main():
    """メイン処理"""
    print("=== EDINET公式英語名辞書の構築 ===\n")

    # 辞書構築
    english_dict, stats = build_english_dict_from_edinet()

    # 統計表示
    print(f"総企業数: {stats['total']}社")
    print(f"英語名なし: {stats['no_english']}社")
    print(f"抽出失敗: {stats['extraction_failed']}社")
    print(f"辞書追加: {stats['added']}件")
    print(f"辞書更新: {stats['updated']}件")
    print(f"重複スキップ: {stats['duplicate_skipped']}件")
    print(f"共通接頭辞による汎用化: {stats['prefix_overrides']}件")
    print(f"\n最終辞書エントリ数: {len(english_dict)}件\n")

    # サンプル表示（アルファベット順）
    print("=== サンプル（最初の30件） ===")
    for i, (eng, kata) in enumerate(sorted(english_dict.items())[:30], 1):
        print(f"{i:3d}. {eng:20s} → {kata}")

    if len(english_dict) > 30:
        print(f"... 他 {len(english_dict) - 30}件")

    # JSONファイルに保存
    output_file = 'english_katakana_dict.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(english_dict, f, ensure_ascii=False, indent=2, sort_keys=True)

    print(f"\n辞書を {output_file} に保存しました")
    print("\nこのファイルは自動的に edinet_cache.py で読み込まれます。")


if __name__ == "__main__":
    main()
