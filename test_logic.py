import csv

dict_items_order = []
prev_item = ""
first_year = ""
current_year = ""

with open('/home/tomo/work_office/download20260223_1843.csv', encoding='cp932', errors='replace') as f:
    reader = csv.reader(f)
    for row in reader:
        if not row or len(row) < 1: continue
        col_0 = row[0].strip()
        
        if "現在" in col_0:
            current_year = col_0.replace("現在", "").strip()
            if first_year == "": first_year = current_year
            continue
            
        if col_0 == "表名称":
            if row[1].strip() == "連結損益（及び包括利益）計算書":
                prev_item = ""
            continue
            
        if "連結損益計算書" in col_0 or "連結損益（及び包括利益）計算書" in col_0:
            if not row[2].strip():
                prev_item = ""
                flag_pl = True
            continue
            
        if col_0 in ["企業名", "証券ｺｰﾄﾞ", "（百万円）"]:
            continue
            
        # simulating the logic
        if current_year in ["2025/03/31", "2024/03/31"]:
            item_name = col_0
            if item_name not in dict_items_order:
                if prev_item in dict_items_order:
                    idx = dict_items_order.index(prev_item)
                    dict_items_order.insert(idx + 1, item_name)
                else:
                    dict_items_order.append(item_name)
            prev_item = item_name

for i, x in enumerate(dict_items_order):
    if x in ['非支配株主に帰属する当期純利益', '営業外収益合計', '支払利息']:
        print(f"{i}: {x}")
print(f"Total: {len(dict_items_order)}")

