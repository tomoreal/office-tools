import os
import sys
import zipfile
import tempfile
import shutil
import glob
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Helper to find specific linkbase/instance files in the unzipped folder
def find_xbrl_files(extract_dir):
    files = {'lab': []}
    base_path = os.path.join(extract_dir, 'XBRL', 'PublicDoc')
    if not os.path.exists(base_path):
        return None
        
    for f in os.listdir(base_path):
        full_path = os.path.join(base_path, f)
        if f.endswith('_pre.xml'):
            files['pre'] = full_path
        elif f.endswith('_lab.xml') and not f.endswith('_lab-en.xml'):
            files['lab'].append(full_path)
        elif f.endswith('.xbrl'):
            files['xbrl'] = full_path
    
    return files if 'pre' in files and 'xbrl' in files else None

import json
import urllib.request
import re

def get_standard_labels(year, cache_dir='edinet_taxonomies'):
    urls = {
        '2025': 'https://www.fsa.go.jp/search/20241112/1c_Taxonomy.zip',
        '2024': 'https://www.fsa.go.jp/search/20231211/1c_Taxonomy.zip',
        '2023': 'https://www.fsa.go.jp/search/20221108/1c_Taxonomy.zip',
        '2022': 'https://www.fsa.go.jp/search/20211109/1c_Taxonomy.zip',
        '2021': 'https://www.fsa.go.jp/search/20201110/1c_Taxonomy.zip',
        '2020': 'https://www.fsa.go.jp/search/20191101/1c_Taxonomy.zip',
        '2019': 'https://www.fsa.go.jp/search/20190228/1c_Taxonomy.zip',
        '2018': 'https://www.fsa.go.jp/search/20180228/1c_Taxonomy.zip',
    }
    
    if year not in urls:
        print(f"Taxonomy for year {year} not found in our known URL map.")
        return {}
        
    tax_dir = os.path.join(cache_dir, year)
    labels_cache_file = os.path.join(tax_dir, 'standard_labels.json')
    
    if os.path.exists(labels_cache_file):
        with open(labels_cache_file, 'r', encoding='utf-8') as f:
            return json.load(f)
            
    if not os.path.exists(tax_dir):
        os.makedirs(tax_dir, exist_ok=True)
        zip_path = os.path.join(tax_dir, 'taxonomy.zip')
        print(f"Downloading EDINET taxonomy for {year} (takes a moment)...")
        try:
            urllib.request.urlretrieve(urls[year], zip_path)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tax_dir)
            os.remove(zip_path)
        except Exception as e:
            print(f"Failed to download/extract taxonomy for {year}: {e}")
            return {}
            
    print(f"Parsing EDINET taxonomy labels for {year}...")
    lab_files = glob.glob(os.path.join(tax_dir, '**', '*_lab.xml'), recursive=True)
    all_labels = {}
    
    for lf in lab_files:
        if 'deprecated' in lf or 'dep' in lf or '-en.xml' in lf:
            continue
        try:
            parsed = parse_labels_file(lf)
            all_labels.update(parsed)
        except:
            pass
            
    if all_labels:
        with open(labels_cache_file, 'w', encoding='utf-8') as f:
            json.dump(all_labels, f, ensure_ascii=False, indent=2)
            
    return all_labels


def parse_labels_file(lab_file):
    labels = {}
    try:
        with open(lab_file, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f.read(), 'lxml')
    except:
        return labels

    locs = soup.find_all('link:loc')
    if not locs: # BeautifulSoup with lxml might lowercase tags
        locs = soup.find_all('link:loc'.lower())
        
    href_to_label_id = {}
    for loc in locs:
        href = loc.get('xlink:href')
        label_id = loc.get('xlink:label')
        if href and label_id:
            element_name = href.split('#')[-1]
            href_to_label_id[label_id] = element_name
            
    arcs = soup.find_all('link:labelarc')
    if not arcs:
        arcs = soup.find_all('link:labelArc')
        
    label_id_to_res_id = {}
    for arc in arcs:
        from_id = arc.get('xlink:from')
        to_id = arc.get('xlink:to')
        if from_id and to_id:
            label_id_to_res_id[from_id] = to_id
            
    res_labels = soup.find_all('link:label')
    if not res_labels:
        res_labels = soup.find_all('link:label'.lower())
        
    res_id_to_text = {}
    for res in res_labels:
        res_id = res.get('xlink:label')
        role = res.get('xlink:role')
        if res.get('xml:lang') == 'ja':
            text = res.text.strip()
            if role == 'http://www.xbrl.org/2003/role/verboseLabel':
                res_id_to_text[res_id] = text 
            elif res_id not in res_id_to_text:
                res_id_to_text[res_id] = text

    for label_id, element_name in href_to_label_id.items():
        res_id = label_id_to_res_id.get(label_id)
        if res_id:
            text = res_id_to_text.get(res_id)
            if text:
                labels[element_name] = text
                if '_' in element_name:
                    base = element_name.split('_')[-1]
                    # Also map the un-prefixed version so dimension parsing can find it
                    labels[base] = text
                
    return labels

def convert_camel_case_to_title(name):
    # e.g. CashAndDeposits -> Cash And Deposits
    import re
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1 \2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1 \2', s1).title()

def build_presentation_tree(pre_file):
    print("Parsing presentation tree...", pre_file)
    with open(pre_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'xml')

    # Group by Role (e.g., BalanceSheet, IncomeStatement)
    roles = {}
    role_refs = soup.find_all('link:roleRef')
    # Actually presentationLinks group by role
    pre_links = soup.find_all('link:presentationLink')
    
    statement_trees = {}
    
    for link in pre_links:
        role = link.get('xlink:role')
        if not role: continue
        
        # Filter roles for important statements (Consolidated BS, PL, CF)
        # Roles look like: http://disclosure.edinet-fsa.go.jp/role/jppfs/rol_BalanceSheet
        role_name = role.split('/')[-1]
        
        # Map locators
        locs = link.find_all('link:loc')
        label_to_element = {}
        for loc in locs:
            href = loc.get('xlink:href')
            label = loc.get('xlink:label')
            if href and label:
                element_name = href.split('#')[-1]
                label_to_element[label] = element_name
                
        # Map arcs (parent-child relationships)
        arcs = link.find_all('link:presentationArc')
        
        parent_child = []
        for arc in arcs:
            from_id = arc.get('xlink:from')
            to_id = arc.get('xlink:to')
            order = arc.get('order')
            if from_id and to_id:
                parent_child.append({
                    'parent': label_to_element.get(from_id),
                    'child': label_to_element.get(to_id),
                    'order': float(order) if order else 0
                })
                
        # Build hierarchy
        if parent_child:
            statement_trees[role_name] = parent_child
            
    return statement_trees

def parse_instance_contexts_and_units(xbrl_file, labels_map):
    print("Parsing XBRL contexts and units...", xbrl_file)
    with open(xbrl_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'xml')
        
    contexts = {}
    
    for ctx in soup.find_all('context'):
        ctx_id = ctx.get('id')
        if not ctx_id:
            continue
            
        members = ctx.find_all('xbrldi:explicitMember')
        dimension_names = []
        for m in members:
            member_val = m.text.split(':')[-1]
            if member_val in labels_map:
                dimension_names.append(labels_map[member_val])
            elif member_val.endswith('Member'):
                # fallback e.g. EnergySegmentMember -> Energy Segment
                dimension_names.append(convert_camel_case_to_title(member_val.replace('Member', '')))
            else:
                dimension_names.append(member_val)
                
        dim_str = "、".join(dimension_names)
        # Clean up verbose XBRL labels
        dim_str = dim_str.replace(' [メンバー]', '').replace('、報告セグメント', '').replace('非連結又は個別', '単体')
            
        period = ctx.find('period')
        if period:
            instant = period.find('instant')
            end_date = period.find('endDate')
            
            p_val = None
            if instant:
                p_val = instant.text
            elif end_date:
                p_val = end_date.text
                
            if p_val:
                contexts[ctx_id] = (p_val, dim_str)
                
    units = {}
    for unit in soup.find_all('unit'):
        unit_id = unit.get('id')
        if not unit_id: continue
        
        is_jpy = False
        if not unit.find('divide'):
            measure = unit.find('measure')
            if measure and 'JPY' in measure.text:
                is_jpy = True
                
        units[unit_id] = is_jpy
                
    return contexts, units

def parse_ixbrl_facts(ixbrl_files, contexts, units):
    facts = []
    for f in ixbrl_files:
        print("Parsing Inline XBRL...", f)
        # using lxml for faster html parsing, large files might take a while
        with open(f, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file.read(), 'lxml')
            
        tags = soup.find_all(['ix:nonfraction', 'ix:nonnumeric'])
        for tag in tags:
            ctx_ref = tag.get('contextref')
            if ctx_ref and ctx_ref in contexts:
                element_name = tag.get('name')
                if element_name and ':' in element_name:
                    # EDINET pre.xml uses jppfs_cor_CashAndDeposits, IXBRL uses jppfs_cor:CashAndDeposits
                    element_name = element_name.replace(':', '_')
                elif element_name:
                    element_name = element_name
                    
                unit_ref = tag.get('unitref')
                is_jpy = units.get(unit_ref, False) if unit_ref else False
                    
                value = tag.text.strip()
                sign = tag.get('sign')
                valStr = ""
                # handle signs, scales etc
                scale = tag.get('scale', '0')
                if tag.name == 'ix:nonfraction' and (value.replace(',','').isdigit() or (value.startswith('-') and value[1:].replace(',','').isdigit())):
                    try:
                        amt_val = value.replace(',', '')
                        amt = float(amt_val)
                        if sign == '-': amt *= -1
                        amt *= (10 ** int(scale))
                        
                        if is_jpy:
                            amt /= 1000000.0
                            
                        valStr = str(int(amt)) if amt.is_integer() else str(amt)
                    except:
                        valStr = value
                else:
                    valStr = value
                    
                facts.append({
                    'element': element_name,
                    'context': ctx_ref,
                    'period': contexts[ctx_ref][0],
                    'dimension': contexts[ctx_ref][1],
                    'value': valStr
                })
    return facts


def create_hierarchy(parent_child_arcs):
    # parent_child_arcs: list of dicts {'parent': p, 'child': c, 'order': o}
    children_map = {}
    roots = set()
    all_children = set()
    
    for arc in parent_child_arcs:
        p = arc['parent']
        c = arc['child']
        if p not in children_map:
            children_map[p] = []
        children_map[p].append((arc['order'], c))
        all_children.add(c)
        roots.add(p)
        
    # True roots are those not in all_children
    true_roots = roots - all_children
    
    for p in children_map:
        children_map[p].sort() # sort by order
        
    ordered_items = []
    
    def traverse(node, path):
        # We store path like A::B::C to handle duplicates
        full_path = "::".join(path + [node])
        ordered_items.append((node, full_path, len(path)))
        if node in children_map:
            for _, child in children_map[node]:
                traverse(child, path + [node])
                
    for r in sorted(list(true_roots)): # sorting roots just in case
        traverse(r, [])
        
    return ordered_items

def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_xbrl_to_excel.py <path_to_zip1> [<path_to_zip2> ...]")
        sys.exit(1)
        
    global_element_period_values = {} # {element: {col_key: value}}
    merged_trees = {} # {role_name: {(parent, child): order}}
    labels_map = {} # {element: label_text}
    labels_map = {} # {element: label_text}
    
    periods_seen = set()

    temp_base = tempfile.mkdtemp()
    
    try:
        # Process each zip file
        for zip_idx, zip_path in enumerate(sys.argv[1:]):
            print(f"Processing {zip_path}...")
            if not os.path.exists(zip_path):
                print(f"File not found: {zip_path}")
                continue
                
            extract_dir = os.path.join(temp_base, f"zip_{zip_idx}")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                
            # EDINET zips have a subdirectory (e.g., S100W178) along with XbrlSearchDlInfo.csv
            subdirs = [d for d in os.listdir(extract_dir) if os.path.isdir(os.path.join(extract_dir, d))]
            if len(subdirs) == 1:
                extract_dir = os.path.join(extract_dir, subdirs[0])
                
            xbrl_files = find_xbrl_files(extract_dir)
            if not xbrl_files:
                print(f"Could not find XBRL files in {zip_path}")
                continue
                
            # Extract taxonomy year from _pre.xml to fetch standard labels
            taxonomy_year = None
            if xbrl_files['pre']:
                with open(xbrl_files['pre'], 'r', encoding='utf-8') as f:
                    content = f.read(4000) # Read start of file
                    m = re.search(r'http://disclosure\.edinet-fsa\.go\.jp/taxonomy/jppfs/(\d{4})-\d{2}-\d{2}', content)
                    if m:
                        taxonomy_year = m.group(1)
            
            if taxonomy_year:
                std_labels = get_standard_labels(taxonomy_year)
                labels_map.update(std_labels)

            # Extract local report labels
            for lf in xbrl_files.get('lab', []):
                print(f"Parsing local labels... {lf}")
                local_labels = parse_labels_file(lf)
                labels_map.update(local_labels)
            
            trees = build_presentation_tree(xbrl_files['pre'])
            
            contexts, units = parse_instance_contexts_and_units(xbrl_files['xbrl'], labels_map)
            
            # Find all inline xbrl files for facts
            ix_files = glob.glob(os.path.join(extract_dir, 'XBRL', 'PublicDoc', '*_ixbrl.htm'))
            facts = parse_ixbrl_facts(ix_files, contexts, units)
            
            # Map facts to elements
            for f in facts:
                el = f['element']
                period = f['period']
                dim = f.get('dimension', '')
                val = f['value']
                
                # Consolidate non-segmented items into ''
                if 'NonConsolidatedMember' in f['context'] or 'NonConsolidated' in dim:
                    pass # We typically want consolidated facts if present
                
                dim_label = dim if dim else "全体"
                col_key = (dim_label, period)
                
                if el not in global_element_period_values:
                    global_element_period_values[el] = {}
                
                # If there are duplicates (e.g. CurrentYearInstant vs CurrentYearDuration with same end date)
                # We overwrite
                global_element_period_values[el][col_key] = val
                periods_seen.add(col_key)
                
            # Merge presentation trees
            for role, tree_arcs in trees.items():
                base_name = role.split('_')[-1]
                # Merge SegmentInformation variants (like Related Info, Goodwill) into a single role
                sub_role_idx = 0
                if 'SegmentInformation' in base_name and '-' in base_name:
                    parts = base_name.rsplit('-', 1)
                    if parts[1].isdigit():
                        sub_role_idx = int(parts[1]) * 1000
                    role = parts[0]
                if not (base_name.startswith('Consolidated') or base_name.startswith('Statement') or base_name.startswith('BalanceSheet') or base_name.startswith('Notes')):
                    continue
                
                if role not in merged_trees:
                    merged_trees[role] = {}
                    
                for arc in tree_arcs:
                    p, c, o = arc['parent'], arc['child'], arc['order']
                    merged_trees[role][(p, c)] = float(o) + sub_role_idx

    finally:
        shutil.rmtree(temp_base)

    all_years_data = {} # {role_name: {hierarchical_key: {period: value}}}
    role_to_order = {} # {role_name: [hierarchical_key1, ...]}
    
    for role, pd_dict in merged_trees.items():
        tree_arcs = [{'parent': p, 'child': c, 'order': o} for (p, c), o in pd_dict.items()]
        ordered_items = create_hierarchy(tree_arcs)
        
        all_years_data[role] = {}
        role_to_order[role] = []
        for el, full_path, depth in ordered_items:
            role_to_order[role].append(full_path)
            all_years_data[role][full_path] = {}
            if el in global_element_period_values:
                for period, val in global_element_period_values[el].items():
                    all_years_data[role][full_path][period] = val

    # Try to find company name for filename
    company_name = "企業名不明"
    name_elements = ['jpcrp_cor_CompanyNameCoverPage', 'jpdei_cor_EntityNameCompanyName']
    for ne in name_elements:
        if ne in global_element_period_values:
            # Get the most recent value if possible
            vals = global_element_period_values[ne]
            if vals:
                # Pick the latest one
                sorted_keys = sorted(vals.keys(), key=lambda x: x[1] if isinstance(x, tuple) else x, reverse=True)
                company_name = vals[sorted_keys[0]]
                break

    # Now generate Excel
    print(f"Generating Excel for {company_name}...")
    wb = Workbook()
    wb.remove(wb.active) # remove default sheet
    
    # Identify periods that are standalone (not consolidated)
    periods_with_standalone = set()
    for role, ordered_keys_dict in all_years_data.items():
        for full_path, p_dict in ordered_keys_dict.items():
            for c in p_dict.keys():
                dim, period = c if isinstance(c, tuple) else ("全体", c)
                if dim == '単体':
                    periods_with_standalone.add(period)
                    
    sorted_periods = sorted(list(periods_seen))
    
    used_sheet_names = set()
    
    for role, ordered_keys in role_to_order.items():
        # Clean role name for sheet
        base_name = role.split('_')[-1]
        sheet_mapping = {
            'ConsolidatedBalanceSheet': '連結貸借対照表',
            'ConsolidatedStatementOfIncome': '連結損益計算書',
            'ConsolidatedStatementOfComprehensiveIncome': '連結包括利益計算書',
            'ConsolidatedStatementOfChangesInEquity': '連結株主資本等変動計算書',
            'ConsolidatedStatementOfChangesInNetAssets': '連結株主資本等変動計算書',
            'ConsolidatedStatementOfCashFlows': '連結キャッシュ・フロー計算書',
            'ConsolidatedStatementOfCashFlows-indirect': '連結キャッシュ・フロー計算書',
            'ConsolidatedStatementOfCashFlows-direct': '連結キャッシュ・フロー計算書',
            'BalanceSheet': '貸借対照表',
            'StatementOfIncome': '損益計算書',
            'StatementOfComprehensiveIncome': '包括利益計算書',
            'StatementOfChangesInEquity': '株主資本等変動計算書',
            'StatementOfChangesInNetAssets': '株主資本等変動計算書',
            'StatementOfCashFlows': 'キャッシュ・フロー計算書',
            'StatementOfCashFlows-indirect': 'キャッシュ・フロー計算書',
            'StatementOfCashFlows-direct': 'キャッシュ・フロー計算書'
        }
        
        japanese_name = sheet_mapping.get(base_name)
        if not japanese_name:
            if base_name.startswith('Notes'):
                sub_name = base_name[5:] # remove 'Notes'
                if sub_name.startswith('SegmentInformation'):
                    m = re.search(r'-(\d+)$', sub_name)
                    segment_dict = {
                        '01': '報告セグメントの概要等',
                        '02': 'セグメント情報',
                        '03': '差異調整事項等',
                        '04': '関連情報',
                        '05': '減損損失',
                        '06': 'のれん',
                        '07': '負ののれん'
                    }
                    if m and m.group(1) in segment_dict:
                        japanese_name = f'注記_{segment_dict[m.group(1)]}'
                    elif m:
                        japanese_name = f'注記_セグメント情報{int(m.group(1))}'
                    else:
                        japanese_name = '注記_セグメント情報'
                else:
                    japanese_name = '注記_' + sheet_mapping.get(sub_name, sub_name)
            else:
                japanese_name = base_name
                
        # In Japanese, 31 characters is usually enough, but we truncate just in case
        sheet_name = japanese_name[:31]
        
        # Collect columns relevant to THIS role based on sheet type
        is_segment = 'セグメント' in sheet_name
        is_consolidated = '連結' in sheet_name
        is_non_consolidated = not is_consolidated and not is_segment and '注記' not in sheet_name
        
        role_columns = set()
        for full_path in ordered_keys:
            if full_path in all_years_data[role]:
                for c in all_years_data[role][full_path].keys():
                    dim, period = c if isinstance(c, tuple) else ("全体", c)
                    
                    if is_segment:
                        if dim == '単体': continue
                    elif is_consolidated:
                        if dim not in ('全体', '連結', '全社'): continue
                    elif is_non_consolidated:
                        if period in periods_with_standalone:
                            if dim != '単体': continue
                        else:
                            if dim not in ('全体', '連結', '全社'): continue
                    else: # other notes (not segment) - keep consolidated
                        if period in periods_with_standalone:
                            if dim == '単体': continue
                            
                    role_columns.add(c)
                    
        if not role_columns:
            continue
            
        counter = 1
        while sheet_name in used_sheet_names:
            suffix = str(counter)
            # Truncate japanese_name enough to fit the suffix and keep total <= 31
            sheet_name = japanese_name[:31 - len(suffix)] + suffix
            counter += 1
            
        used_sheet_names.add(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        
        # Track separators so we only print them once per sheet
        seen_related = False
        seen_goodwill = False
        seen_negative_goodwill = False
        seen_impairment = False
        
        # Sort columns logically (Segment first: 全体->各セグメント->調整額->合計, then dates)
        def sort_col(c):
            if isinstance(c, str):
                return (0, "", c)
            dim, period = c
            
            order = 10
            if dim in ('全体', '連結', '全社'):
                order = 0
            elif dim == '単体':
                order = 1
            elif '報告セグメント以外' in dim or 'その他' in dim:
                order = 90
            elif dim in ('調整額', '全社・消去', '消去又は全社'):
                order = 98
            elif dim in ('合計', '連結財務諸表計上額'):
                order = 99
                
            return (order, dim, period)
            
        sorted_role_cols = sorted(list(role_columns), key=sort_col)
        
        has_segments = any(isinstance(c, tuple) and c[0] not in ('全体', '連結', '全社') for c in sorted_role_cols)
        
        if has_segments:
            # Two-tier header: Row 1 = Segments, Row 2 = Dates
            headers_row1 = [""] + [c[0] if isinstance(c, tuple) else "" for c in sorted_role_cols]
            headers_row2 = ["勘定科目"] + [c[1] if isinstance(c, tuple) else c for c in sorted_role_cols]
            ws.append(headers_row1)
            ws.append(headers_row2)
        else:
            headers = ["勘定科目"] + [c[1] if isinstance(c, tuple) else c for c in sorted_role_cols]
            ws.append(headers)
        
        for full_path in ordered_keys:
            # Extract element name to get label
            el = full_path.split('::')[-1]
            
            # Common terminology translations as a fallback
            common_dict = {
                'CashAndDeposits': '現金及び預金',
                'NotesAndAccountsReceivableTrade': '受取手形及び売掛金',
                'MerchandiseAndFinishedGoods': '商品及び製品',
                'RawMaterialsAndSupplies': '原材料及び貯蔵品',
                'CurrentAssets': '流動資産',
                'NoncurrentAssets': '固定資産',
                'PropertyPlantAndEquipment': '有形固定資産',
                'IntangibleAssets': '無形固定資産',
                'InvestmentsAndOtherAssets': '投資その他の資産',
                'Assets': '資産合計',
                'CurrentLiabilities': '流動負債',
                'NoncurrentLiabilities': '固定負債',
                'Liabilities': '負債合計',
                'NetAssets': '純資産合計',
                'LiabilitiesAndNetAssets': '負債純資産合計',
                'NetSales': '売上高',
                'CostOfSales': '売上原価',
                'GrossProfit': '売上総利益',
                'SellingGeneralAndAdministrativeExpenses': '販売費及び一般管理費',
                'OperatingIncome': '営業利益',
                'OrdinaryIncome': '経常利益',
                'NetIncome': '当期純利益'
            }
            
            label = labels_map.get(el)
            if not label:
                parts = el.split('_')
                base_name = parts[-1] if len(parts) > 1 else el
                if base_name in common_dict:
                    label = common_dict[base_name]
                else:
                    label = convert_camel_case_to_title(base_name)
                    
            # Insert Separator rows for Segment Notes subsets
            if 'SegmentInformation' in role:
                if not seen_related and ('NotesInformationAssociated' in el or 'DisclosureOfRelatedInformation' in el):
                    ws.append([])
                    ws.append(["【 注記：関連情報 】"])
                    seen_related = True
                elif 'Goodwill' in el and ('Amortization' in el or 'Negative' in el or 'Disclosure' in el):
                    if 'Negative' in el and not seen_negative_goodwill:
                        ws.append([])
                        ws.append(["【 注記：負ののれん 】"])
                        seen_negative_goodwill = True
                    elif 'Negative' not in el and not seen_goodwill:
                        ws.append([])
                        ws.append(["【 注記：のれん 】"])
                        seen_goodwill = True
                elif not seen_impairment and ('ImpairmentLoss' in el and 'Segment' in el):
                    ws.append([])
                    ws.append(["【 注記：減損損失 】"])
                    seen_impairment = True
                    
            # Indent based on depth? Depth is full_path count
            depth = len(full_path.split('::')) - 1
            indent_prefix = "　" * depth
            
            row_data = [indent_prefix + label]
            
            has_data = False
            for col_key in sorted_role_cols:
                val = all_years_data[role][full_path].get(col_key, "")
                # Clean numeric values
                if val and not val.isalpha():
                    try:
                        val = float(val)
                    except:
                        pass
                row_data.append(val)
                if val != "":
                    has_data = True
                    
            if has_data: # Only append rows that have at least one value across columns
                ws.append(row_data)

        # Apply formatting and column widths
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    # Format: #,##0_;[Red]-#,##0
                    cell.number_format = '#,##0_ ;[Red]\-#,##0 '
                    
        # Auto-adjust column widths
    for out_ws in wb.worksheets:
        for col in out_ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            # Add a little extra padding, especially for Japanese characters
            adjusted_width = (max_length + 2) * 1.2
            # Cap width to prevent massive columns from long text
            if adjusted_width > 50:
                adjusted_width = 50
            out_ws.column_dimensions[column_letter].width = adjusted_width

    out_file = f'XBRL_横展開_{company_name}.xlsx'
    wb.save(out_file)
    print(f"Saved to {out_file}")


if __name__ == "__main__":
    main()
