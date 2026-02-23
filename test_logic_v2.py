import csv

dict_items_order = []
prev_item = ""
first_year = ""
current_year = ""
current_type = ""

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
            current_type = row[1].strip()
            prev_item = ""
            continue
            
        if "連結" in col_0 or "計算書" in col_0:
            if len(row) > 2 and not row[2].strip():
                current_type = col_0
                prev_item = ""
            continue
            
        if col_0 in ["企業名", "証券ｺｰﾄﾞ", "（百万円）"] or ("/" in col_0 and "-" in col_0):
            continue
            
        if current_type == "連結損益計算書" or current_type == "連結損益（及び包括利益）計算書":
            item_name = col_0
            if item_name not in dict_items_order:
                if current_year != first_year and prev_item in dict_items_order:
                    idx = dict_items_order.index(prev_item)
                    dict_items_order.insert(idx + 1, item_name)
                else:
                    dict_items_order.append(item_name)
            prev_item = item_name

for i, x in enumerate(dict_items_order):
    if "純利益" in x or "営業外" in x or "支払利息" in x:
        print(f"{i}: {x}")

print(dict_items_order[-10:])
