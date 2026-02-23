import zipfile
import xml.etree.ElementTree as ET
import sys

try:
    path = '/home/tomo/work_office/zaimu_data_dl_20260223_1832.xlsx'
    zf = zipfile.ZipFile(path)
    
    # 共有文字列の取得
    shared_strings = []
    try:
        ss_data = zf.read('xl/sharedStrings.xml')
        ss_root = ET.fromstring(ss_data)
        # prefixなしだと探しにくいので雑にタグ名で検索
        for t in ss_root.iter():
            if t.tag.endswith('}t'):
                shared_strings.append(t.text if t.text else "")
            # fallback for empty strings, text might be in <r><t>
    except KeyError as e:
        print("No sharedStrings:", e)
    
    # シート1の取得
    sheet_data = zf.read('xl/worksheets/sheet1.xml')
    sheet_root = ET.fromstring(sheet_data)
    
    print("--- Sheet 1 first 50 rows ---")
    count = 0
    for row in sheet_root.iter():
        if row.tag.endswith('}row'):
            row_data = []
            for c in row.iter():
                if c.tag.endswith('}c'):
                    val = ""
                    t_attr = c.attrib.get('t')
                    # <v> tag contains value
                    for v in c.iter():
                        if v.tag.endswith('}v'):
                            val = v.text
                    if t_attr == 's' and val and val.isdigit():
                        idx = int(val)
                        if idx < len(shared_strings):
                            val = shared_strings[idx]
                    row_data.append(val)
            # clean up duplicate v parsing
            # simplify printing
            if row_data:
                print(f"Row {count+1}: {row_data}")
            count += 1
            if count >= 50:
                break
                
except Exception as e:
    print("Error:", e)
