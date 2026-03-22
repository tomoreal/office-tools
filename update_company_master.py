#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINETのコードリスト（EdinetcodeDlInfo.csv）から、
edinet_cache.db に company_master テーブルを作成・更新するスクリプト。
これにより英語名・カタカナ名での直接検索が可能になる。
"""

import os
import csv
import sqlite3

def update_company_master(csv_path: str = 'EdinetcodeDlInfo.csv', db_path: str = 'edinet_cache.db'):
    if not os.path.exists(csv_path):
        print(f"エラー: {csv_path} が見つかりません。")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # テーブルの作成
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS company_master (
            edinet_code TEXT PRIMARY KEY,
            japanese_name TEXT,
            english_name TEXT,
            kana_name TEXT
        )
    """)
    
    # 検索用インデックス
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_cm_english 
        ON company_master(english_name)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_cm_kana 
        ON company_master(kana_name)
    """)

    print("CSVデータを読み込んでいます...")
    
    count = 0
    # 一括処理用のリスト
    records = []
    
    with open(csv_path, 'r', encoding='cp932', errors='ignore') as f:
        reader = csv.reader(f)
        
        # ヘッダー2行をスキップ
        next(reader, None)
        next(reader, None)
        
        for row in reader:
            if len(row) < 9:
                continue
                
            edinet_code = row[0].strip()
            if not edinet_code.startswith('E') and not edinet_code.startswith('G'):
                # EかGから始まるEDINETコード以外はスキップ（念のため）
                pass
                
            japanese_name = row[6].strip()
            english_name = row[7].strip()
            kana_name = row[8].strip()
            
            records.append((edinet_code, japanese_name, english_name, kana_name))
            count += 1

    print(f"{count}件のデータをデータベースに挿入しています...")
    
    # バルクインサート
    cursor.executemany("""
        INSERT OR REPLACE INTO company_master 
        (edinet_code, japanese_name, english_name, kana_name) 
        VALUES (?, ?, ?, ?)
    """, records)

    conn.commit()
    conn.close()
    
    print("完了しました。")

if __name__ == "__main__":
    update_company_master()
