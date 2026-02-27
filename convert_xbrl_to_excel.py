import os
import sys
import zipfile
import tempfile
import shutil
import glob
from bs4 import BeautifulSoup
try:
    from lxml import etree
except ImportError:
    import xml.etree.ElementTree as etree

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
except ImportError:
    Workbook = None

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
    """Returns (all_labels, label_priorities) for the given taxonomy year.
    Uses cached standard_labels.json if it exists.
    """
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
        print(f"Taxonomy for year {year} not found in our known URL map.", file=sys.stderr)
        return {}, {}
        
    tax_dir = os.path.join(cache_dir, str(year))
    labels_cache_file = os.path.join(tax_dir, 'standard_labels.json')
    
    if os.path.exists(labels_cache_file):
        with open(labels_cache_file, 'r', encoding='utf-8') as f:
            labels = json.load(f)
            # Default priority for cached labels: 50 (middle ground)
            priorities = {k: 50 for k in labels}
            return labels, priorities
            
    if not os.path.exists(tax_dir):
        os.makedirs(tax_dir, exist_ok=True)
        zip_path = os.path.join(tax_dir, 'taxonomy.zip')
        print(f"Downloading EDINET taxonomy for {year} (takes a moment)...", file=sys.stderr)
        try:
            urllib.request.urlretrieve(urls[year], zip_path)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tax_dir)
            os.remove(zip_path)
        except Exception as e:
            print(f"Failed to download/extract taxonomy for {year}: {e}", file=sys.stderr)
            return {}, {}
            
    print(f"Parsing EDINET taxonomy labels for {year}...", file=sys.stderr)
    lab_files = glob.glob(os.path.join(tax_dir, '**', '*_lab.xml'), recursive=True)
    all_labels = {}
    label_priorities = {} # {element_name: priority}
    
    for lf in lab_files:
        if 'deprecated' in lf or 'dep' in lf or '-en.xml' in lf:
            continue
        try:
            parsed_labels, parsed_priorities = parse_labels_file(lf)
            for el, text in parsed_labels.items():
                prio = parsed_priorities.get(el, 99)
                if el not in all_labels or prio < label_priorities.get(el, 100):
                    all_labels[el] = text
                    label_priorities[el] = prio
        except:
            pass
            
    if all_labels:
        with open(labels_cache_file, 'w', encoding='utf-8') as f:
            json.dump(all_labels, f, ensure_ascii=False, indent=2)
            
    return all_labels, label_priorities


def parse_labels_file(lab_file):
    """Parse an XBRL label linkbase using lxml for robust namespace handling.
    Returns (labels, priorities) where labels is a dict mapping element names to text,
    and priorities maps them to their best priority score.
    """
    labels = {}
    priorities = {}
    try:
        # Parse XML with lxml to respect namespaces and recover from minor errors
        parser = etree.XMLParser(recover=True)
        tree = etree.parse(lab_file, parser)
    except Exception as e:
        # If parsing fails, return empty mappings
        return labels, priorities

    # Namespace map for XBRL linkbase
    ns = {
        "link": "http://www.xbrl.org/2003/linkbase",
        "xlink": "http://www.w3.org/1999/xlink",
        "xml": "http://www.w3.org/XML/1998/namespace"
    }

    # 1. Locate all <link:loc> elements to map label IDs to element QNames
    href_to_label_id = {}
    for loc in tree.xpath("//link:loc", namespaces=ns):
        href = loc.get("{http://www.w3.org/1999/xlink}href")
        label_id = loc.get("{http://www.w3.org/1999/xlink}label")
        if href and label_id:
            # Element name may be a QName like jppfs_cor:CashAndDeposits
            element_name = href.split('#')[-1].replace(':', '_')
            href_to_label_id[label_id] = element_name

    # 2. Filter arcs to only concept‑label relationships (collect ALL associated resource IDs)
    label_id_to_res_ids = {}
    arc_xpath = "//link:labelArc[@xlink:arcrole='http://www.xbrl.org/2003/arcrole/concept-label']"
    for arc in tree.xpath(arc_xpath, namespaces=ns):
        from_id = arc.get("{http://www.w3.org/1999/xlink}from")
        to_id = arc.get("{http://www.w3.org/1999/xlink}to")
        if from_id and to_id:
            if from_id not in label_id_to_res_ids:
                label_id_to_res_ids[from_id] = []
            label_id_to_res_ids[from_id].append(to_id)

    # 3. Gather label resources (<link:label>) with Japanese language
    res_id_to_text = {}
    res_id_to_priority = {}
    
    # Role priority: verboseLabel is usually best for Excel reports.
    # We explicitly lower priority for generic labels like "合計" or "計".
    # XBRL Label Roles and their associated priority (lower is better)
    role_priority = {
        "http://www.xbrl.org/2003/role/verboseLabel": 1,
        "http://www.xbrl.org/2003/role/label": 2,
        "http://disclosure.edinet-fsa.go.jp/jpcrp/alt/role/label": 3, # EDINET alternate label
        "http://www.xbrl.org/2003/role/terseLabel": 5,
        "http://www.xbrl.org/2003/role/totalLabel": 10, # Standard total
        "http://disclosure.edinet-fsa.go.jp/jpcrp/alt/role/totalLabel": 11, # EDINET alternate total
    }

    GENERIC_LABELS = ('合計', '計', 'total', 'sum', 'subtotal', '金額')

    for res in tree.xpath("//link:label", namespaces=ns):
        lang = res.get("{http://www.w3.org/XML/1998/namespace}lang")
        if not lang or not lang.startswith('ja'):
            continue
        res_id = res.get("{http://www.w3.org/1999/xlink}label")
        role = res.get("{http://www.w3.org/1999/xlink}role")
        if not res_id:
            continue
            
        text = ''.join(res.itertext()).strip()
        if not text:
            continue
            
        priority = role_priority.get(role, 99)
        # Penalize generic labels to avoid "Total" appearing everywhere if a better name exists
        # 【修正】ペナルティを強化 (20 -> 50)
        if any(g in text.lower() for g in GENERIC_LABELS):
            priority += 50
            
        if (res_id not in res_id_to_text) or (priority < res_id_to_priority.get(res_id, 100)):
            res_id_to_text[res_id] = text
            res_id_to_priority[res_id] = priority

    # 4. Build final mapping (pick the best label text among all resource IDs)
    for label_id, element_name in href_to_label_id.items():
        res_ids = label_id_to_res_ids.get(label_id, [])
        best_text = None
        best_priority = 100
        
        for res_id in res_ids:
            text = res_id_to_text.get(res_id)
            priority = res_id_to_priority.get(res_id, 99)
            if text and priority < best_priority:
                best_text = text
                best_priority = priority
                
        if best_text:
            if element_name not in labels or best_priority < priorities.get(element_name, 100):
                labels[element_name] = best_text
                priorities[element_name] = best_priority
            # [REMOVED] mapping base name to labels[base] as it causes collisions
    return labels, priorities

def convert_camel_case_to_title(name):
    # e.g. CashAndDeposits -> Cash And Deposits
    import re
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1 \2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1 \2', s1).title()

def parse_presentation_linkbase(pre_file):
    print("Parsing presentation linkbase...", pre_file, file=sys.stderr)
    try:
        # Use lxml for robust namespace handling and performance
        parser = etree.XMLParser(recover=True)
        tree = etree.parse(pre_file, parser)
    except Exception as e:
        print(f"Error parsing presentation linkbase: {e}", file=sys.stderr)
        return {}

    ns = {
        "link": "http://www.xbrl.org/2003/linkbase",
        "xlink": "http://www.w3.org/1999/xlink"
    }

    # 1. Group by role URI first
    role_to_content = {} # {role_uri: {'locs': {label: element}, 'arcs': [arc_dicts]}}
    
    links = tree.xpath("//link:presentationLink", namespaces=ns)
    for link in links:
        role_uri = link.get("{http://www.w3.org/1999/xlink}role")
        if not role_uri:
            continue
        
        if role_uri not in role_to_content:
            role_to_content[role_uri] = {'locs': {}, 'arcs': []}
            
        # Map locators in this link
        locs = link.xpath("link:loc", namespaces=ns)
        for loc in locs:
            href = loc.get("{http://www.w3.org/1999/xlink}href")
            label = loc.get("{http://www.w3.org/1999/xlink}label")
            if href and label:
                # Normalize element name: replace ':' with '_' to match facts and labels
                element_name = href.split('#')[-1].replace(':', '_')
                role_to_content[role_uri]['locs'][label] = element_name
                
        # Map arcs in this link
        arcs = link.xpath("link:presentationArc", namespaces=ns)
        for arc in arcs:
            from_id = arc.get("{http://www.w3.org/1999/xlink}from")
            to_id = arc.get("{http://www.w3.org/1999/xlink}to")
            order = arc.get("order")
            if from_id and to_id:
                role_to_content[role_uri]['arcs'].append({
                    'from': from_id,
                    'to': to_id,
                    'order': float(order) if order else 0.0
                })

    statement_trees = {} # {role_name: [arc_dicts]}
    
    for role_uri, content in role_to_content.items():
        role_name = role_uri.split('/')[-1]
        label_to_element = content['locs']
        
        parent_child = []
        for arc in content['arcs']:
            p = label_to_element.get(arc['from'])
            c = label_to_element.get(arc['to'])
            if p and c:
                parent_child.append({
                    'parent': p,
                    'child': c,
                    'order': arc['order']
                })
                
        if not parent_child:
            continue
        
        # Original role
        statement_trees[role_name] = parent_child
        
        # Special logic for "jumbo" roles (e.g. Cabinet Office Ordinance Form 3)
        # These roles often contain many independent financial statements under specific Heading elements.
        jumbo_indicators = ['formno3', 'cabinetofficeordinance', 'annualsecuritiesreport']
        if any(ji in role_name.lower() for ji in jumbo_indicators):
            major_headings = [
                'ConsolidatedBalanceSheetHeading', 'ConsolidatedStatementOfIncomeHeading', 
                'ConsolidatedStatementOfCashFlowsHeading', 'ConsolidatedStatementOfChangesInEquityHeading',
                'BalanceSheetHeading', 'StatementOfIncomeHeading', 'StatementOfChangesInEquityHeading',
                'ConsolidatedStatementOfFinancialPositionIFRSHeading', 'ConsolidatedStatementOfProfitOrLossIFRSHeading',
                'ConsolidatedStatementOfCashFlowsIFRSHeading', 'ConsolidatedStatementOfChangesInEquityIFRSHeading',
                'ConsolidatedStatementOfFinancialPositionHeading', 
                'ConsolidatedStatementOfProfitOrLossHeading',
                'ConsolidatedStatementOfComprehensiveIncomeIFRSHeading',
                'SummaryOfBusinessResultsHeading', 'BusinessResultsOfGroupHeading', 'BusinessResultsOfReportingCompanyHeading'
            ]
            
            all_elements = label_to_element.values()
            for h in major_headings:
                # Look for ALL elements that end with the heading name (handles prefixes and underscores)
                h_elements = [el for el in all_elements if el.endswith(h)]
                
                for h_element in h_elements:
                    # Extract subtree starting from this heading
                    subtree_arcs = []
                    queue = [h_element]
                    seen = {h_element}
                    while queue:
                        curr_parent = queue.pop(0)
                        for arc in parent_child:
                            if arc['parent'] == curr_parent:
                                subtree_arcs.append(arc)
                                if arc['child'] not in seen:
                                    seen.add(arc['child'])
                                    queue.append(arc['child'])
                    
                    if subtree_arcs:
                        virtual_role = h_element
                        if virtual_role.endswith('Heading'):
                            virtual_role = virtual_role[:-7]
                        statement_trees[virtual_role] = subtree_arcs
                        
    return statement_trees

def parse_instance_contexts_and_units(xbrl_file, labels_map):
    print("Parsing XBRL contexts and units...", xbrl_file, file=sys.stderr)
    try:
        # Use lxml for robust namespace handling
        parser = etree.XMLParser(recover=True)
        tree = etree.parse(xbrl_file, parser)
    except Exception as e:
        print(f"Error parsing XBRL instance: {e}", file=sys.stderr)
        return {}, {}

    # Standard namespaces for XBRL instance and dimensions
    ns = {
        "xbrli": "http://www.xbrl.org/2003/instance",
        "xbrldi": "http://xbrl.org/2006/xbrldi"
    }

    contexts = {}
    
    # 1. Parse contexts
    for ctx in tree.xpath("//xbrli:context", namespaces=ns):
        ctx_id = ctx.get('id')
        if not ctx_id:
            continue
            
        members = ctx.xpath(".//xbrldi:explicitMember", namespaces=ns)
        dimension_names = []
        for m in members:
            # Handle QNames in member text (e.g., jppfs_cor:EnergySegmentMember)
            m_text = m.text or ""
            # Try to look up the label using various common prefixes
            # Since we removed the aggressive base-name mapping to fix IFRS labels,
            # we must be more explicit here.
            member_val = m_text.split(':')[-1]
            prefixes = ['jppfs_cor_', 'jpigp_cor_', 'jpcrp_cor_', 'jpdei_cor_', '']
            label = None
            for p in prefixes:
                if p + member_val in labels_map:
                    label = labels_map[p + member_val]
                    break
            
            if label:
                dimension_names.append(label)
            elif member_val.endswith('Member'):
                # Fallback: EnergySegmentMember -> Energy Segment
                dimension_names.append(convert_camel_case_to_title(member_val.replace('Member', '')))
            else:
                dimension_names.append(member_val)
                
        dim_str = "、".join(dimension_names)
        # Clean up verbose XBRL labels
        dim_str = dim_str.replace(' [メンバー]', '').replace('、報告セグメント', '').replace('非連結又は個別', '単体')
            
        period_elem = ctx.xpath("xbrli:period", namespaces=ns)
        if period_elem:
            period_elem = period_elem[0]
            instant = period_elem.xpath("xbrli:instant", namespaces=ns)
            end_date = period_elem.xpath("xbrli:endDate", namespaces=ns)
            
            p_val = None
            if instant:
                p_val = instant[0].text
            elif end_date:
                p_val = end_date[0].text
                
            if p_val:
                contexts[ctx_id] = (p_val, dim_str)
                
    units = {}
    # 2. Parse units
    for unit in tree.xpath("//xbrli:unit", namespaces=ns):
        unit_id = unit.get('id')
        if not unit_id:
            continue
        
        is_jpy = False
        # Only consider simple units (non‑divide) for JPY amount identification
        if not unit.xpath("xbrli:divide", namespaces=ns):
            measure = unit.xpath(".//xbrli:measure", namespaces=ns)
            if measure and 'JPY' in (measure[0].text or ""):
                is_jpy = True
                
        units[unit_id] = is_jpy
                
    return contexts, units

def parse_ixbrl_facts(ixbrl_files, contexts, units):
    facts = []
    
    def get_attr(tag, attr_name):
        for k, v in tag.attrs.items():
            if k.lower() == attr_name.lower():
                return v
        return None

    for f in ixbrl_files:
        print(f"Parsing Inline XBRL... {os.path.basename(f)}", file=sys.stderr)
        with open(f, 'r', encoding='utf-8', errors='replace') as file:
            content = file.read()
            # Try lxml first, fallback to html.parser
            try:
                soup = BeautifulSoup(content, 'lxml')
            except:
                soup = BeautifulSoup(content, 'html.parser')
            
        def is_ix_tag(tag):
            if not tag.name: return False
            local = tag.name.split(':')[-1].lower()
            return local in ('nonfraction', 'nonnumeric')
            
        tags = soup.find_all(is_ix_tag)
        print(f"  Found {len(tags)} tags", file=sys.stderr)
        
        elem_order = 0
        for tag in tags:
            ctx_ref = get_attr(tag, 'contextRef')
            if ctx_ref and ctx_ref in contexts:
                element_name = get_attr(tag, 'name')
                if not element_name: continue
                if ':' in element_name:
                    element_name = element_name.replace(':', '_')
                    
                value = tag.text.strip()
                local_name = tag.name.split(':')[-1].lower()
                
                if local_name == 'nonnumeric':
                    if 'TextBlock' in element_name or len(value) > 200 or '。' in value or '\n' in value:
                        continue
                
                valStr = ""
                if local_name == 'nonfraction':
                    unit_ref = get_attr(tag, 'unitRef')
                    is_jpy = units.get(unit_ref, False) if unit_ref else False
                    sign = get_attr(tag, 'sign')
                    scale = get_attr(tag, 'scale') or '0'
                    
                    clean_val = value.replace(',', '').replace('△', '-').replace('▲', '-').replace('(', '-').replace(')', '').strip()
                    
                    try:
                        amt = float(clean_val)
                        if sign == '-': amt *= -1
                        amt *= (10 ** int(scale))
                        if is_jpy: amt /= 1000000.0
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
                    'value': valStr,
                    'source_file': f,
                    'elem_order': elem_order
                })
                elem_order += 1
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

def process_xbrl_zips(zip_paths, output_dir=None):
    if not zip_paths:
        return None
        
    global_element_period_values = {} # {element: {col_key: value}}
    merged_trees = {} # {role_name: {(parent, child): order}}
    labels_map = {} # {element: label_text}
    labels_map_priorities = {} # {element: priority}
    
    periods_seen = set()
    all_facts = []  # Accumulate facts across all zips for fallback logic

    # Use provided output_dir for temp files if possible to avoid permission issues in /tmp
    parent_temp_dir = output_dir if output_dir and os.path.exists(output_dir) else None
    temp_base = tempfile.mkdtemp(dir=parent_temp_dir)
    
    try:
        # Process each zip file
        for zip_idx, zip_path in enumerate(zip_paths):
            print(f"Processing {zip_path}...", file=sys.stderr)
            if not os.path.exists(zip_path):
                print(f"File not found: {zip_path}", file=sys.stderr)
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
                print(f"Could not find XBRL files in {zip_path}", file=sys.stderr)
                continue
                
            # Extract taxonomy year from _pre.xml to fetch standard labels
            taxonomy_year = None
            if xbrl_files['pre']:
                with open(xbrl_files['pre'], 'r', encoding='utf-8') as f:
                    content = f.read(4000) # Read start of file
                    # Support various standards like jppfs (J-GAAP), jpigp (IFRS), etc.
                    m = re.search(r'http://disclosure\.edinet-fsa\.go\.jp/taxonomy/[a-z]+(?:_[a-z]+)?/(\d{4})-\d{2}-\d{2}', content)
                    if m:
                        taxonomy_year = m.group(1)
            
            if taxonomy_year:
                std_labels, std_priorities = get_standard_labels(taxonomy_year)
                labels_map.update(std_labels)
                labels_map_priorities.update(std_priorities)

            # Extract local report labels
            for lf in xbrl_files.get('lab', []):
                print(f"Parsing local labels... {lf}", file=sys.stderr)
                local_labels, local_priorities = parse_labels_file(lf)
                # Apply priority logic for local labels too
                # Give local labels a slight boost (subtract 1 from priority) to prefer them over standard
                for k, v in local_labels.items():
                    p = local_priorities.get(k, 99) - 1
                    if k not in labels_map or p < labels_map_priorities.get(k, 100):
                        labels_map[k] = v
                        labels_map_priorities[k] = p
            
            trees = parse_presentation_linkbase(xbrl_files['pre'])
            
            contexts, units = parse_instance_contexts_and_units(xbrl_files['xbrl'], labels_map)
            
            # Find all inline xbrl files for facts
            ix_files = glob.glob(os.path.join(extract_dir, 'XBRL', 'PublicDoc', '*_ixbrl.htm'))
            facts = parse_ixbrl_facts(ix_files, contexts, units)
            all_facts.extend(facts)
            
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
                if not (base_name.startswith('Consolidated') or base_name.startswith('Statement') or base_name.startswith('BalanceSheet') or base_name.startswith('Notes') or 'BusinessResults' in base_name):
                    continue
                
                if role not in merged_trees:
                    merged_trees[role] = {}
                    
                for arc in tree_arcs:
                    p, c, o = arc['parent'], arc['child'], arc['order']
                    merged_trees[role][(p, c)] = float(o) + sub_role_idx

    finally:
        shutil.rmtree(temp_base)

    # --- Fallback for old EDINET format (e.g. 2016-2018) ---
    # build synthetic roles from the element appearance order in the known ixbrl files.
    EDINET_DOC_ROLE_MAP = {
        '0105010': 'rol_ConsolidatedBalanceSheet',
        '0105020': 'rol_ConsolidatedStatementOfIncome',
        '0105025': 'rol_ConsolidatedStatementOfComprehensiveIncome',
        '0105040': 'rol_ConsolidatedStatementOfChangesInNetAssets',
        '0105050': 'rol_ConsolidatedStatementOfCashFlows',
    }

    # Correct way to find roles that actually have a presentation structure in merged_trees.
    # We only skip the fallback if at least one zip provided a non-empty tree for that role base name.
    roles_with_structure = set()
    for role_uri_name, arcs_dict in merged_trees.items():
        if len(arcs_dict) > 0:
            roles_with_structure.add(role_uri_name.split('_')[-1])

    roles_to_fill = {
        doc_code: role_name
        for doc_code, role_name in EDINET_DOC_ROLE_MAP.items()
        if role_name.split('_')[-1] not in roles_with_structure
    }

    if roles_to_fill:
        print(f"Applying fallback for missing roles: {list(roles_to_fill.values())}", file=sys.stderr)
        facts_by_doc = {}  # {doc_code: {element: min_order}}
        for f in all_facts:
            src = f.get('source_file', '')
            fname = os.path.basename(src)
            for doc_code in roles_to_fill:
                # Use regex for more reliable document code matching at start of filename
                if re.match(r'^' + doc_code, fname):
                    if doc_code not in facts_by_doc:
                        facts_by_doc[doc_code] = {}
                    el = f['element']
                    order = f.get('elem_order', 0)
                    if el not in facts_by_doc[doc_code] or order < facts_by_doc[doc_code][el]:
                        facts_by_doc[doc_code][el] = order

        for doc_code, role_name in roles_to_fill.items():
            if doc_code not in facts_by_doc:
                continue
            elem_order_map = facts_by_doc[doc_code]
            if not elem_order_map:
                continue

            sorted_elems = sorted(elem_order_map.items(), key=lambda x: x[1])
            virtual_root = role_name
            arcs = []
            for elem, _order in sorted_elems:
                arcs.append({'parent': virtual_root, 'child': elem, 'order': float(_order)})

            if arcs:
                merged_trees[role_name] = {(a['parent'], a['child']): a['order'] for a in arcs}
                print(f"[Fallback] Created synthetic role {role_name} from {doc_code} with {len(arcs)} elements", file=sys.stderr)

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
    print(f"Generating Excel for {company_name}...", file=sys.stderr)
    wb = Workbook()
    default_sheet_removed = False
    
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
            'ConsolidatedStatementOfFinancialPositionIFRS': '連結貸借対照表',
            'ConsolidatedStatementOfProfitOrLossIFRS': '連結損益計算書',
            'ConsolidatedStatementOfComprehensiveIncomeIFRS': '連結包括利益計算書',
            'ConsolidatedStatementOfChangesInEquityIFRS': '連結株主資本等変動計算書',
            'ConsolidatedStatementOfCashFlowsIFRS': '連結キャッシュ・フロー計算書',
            'BalanceSheet': '貸借対照表',
            'StatementOfIncome': '損益計算書',
            'StatementOfComprehensiveIncome': '包括利益計算書',
            'StatementOfChangesInEquity': '株主資本等変動計算書',
            'StatementOfChangesInNetAssets': '株主資本等変動計算書',
            'StatementOfCashFlows': 'キャッシュ・フロー計算書',
            'StatementOfCashFlows-indirect': 'キャッシュ・フロー計算書',
            'StatementOfCashFlows-direct': 'キャッシュ・フロー計算書',
            'SummaryOfBusinessResults': '主要な経営指標等の推移',
            'BusinessResultsOfGroup': '主要な経営指標等の推移（連結）',
            'BusinessResultsOfReportingCompany': '主要な経営指標等の推移（単体）'
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
                
        # Add accounting standard suffix
        is_ifrs = 'IFRS' in base_name
        suffix = '(IFRS)' if is_ifrs else '(日本基準)'
        japanese_name += suffix
        
        # In Japanese, 31 characters maximum for sheet name
        if len(japanese_name) > 31:
            allowed_len = 31 - len(suffix)
            sheet_name = japanese_name[:allowed_len] + suffix
        else:
            sheet_name = japanese_name
        
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
        if not default_sheet_removed:
            wb.remove(wb['Sheet'])
            default_sheet_removed = True
        
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
            headers_row1 = ["", ""] + [c[0] if isinstance(c, tuple) else "" for c in sorted_role_cols]
            headers_row2 = ["勘定科目", "項目（英名）"] + [c[1] if isinstance(c, tuple) else c for c in sorted_role_cols]
            ws.append(headers_row1)
            ws.append(headers_row2)
        else:
            headers = ["勘定科目", "項目（英名）"] + [c[1] if isinstance(c, tuple) else c for c in sorted_role_cols]
            ws.append(headers)
        
        for full_path in ordered_keys:
            # Extract element name to get label
            el = full_path.split('::')[-1]
            
            # Common terminology translations as a fallback
            common_dict = {
                'CashAndDeposits': '現金及び預金',
                'NotesAndAccountsReceivableTrade': '受取手形及び売掛金',
                'MerchandiseAndFinishedGoods': '商品及び製品',
                'Notes': '注記',
                'Inventory': '棚卸資産',
                'Inventories': '棚卸資産',
                'PropertyPlantAndEquipment': '有形固定資産',
                'IntangibleAssets': '無形固定資産',
                'ConsolidatedBalanceSheet': '連結貸借対照表',
                'ConsolidatedStatementOfIncome': '連結損益計算書',
                'ConsolidatedStatementOfCashFlows': '連結キャッシュ・フロー計算書',
                'ConsolidatedStatementOfChangesInEquity': '連結株主資本等変動計算書',
                'ConsolidatedStatementOfFinancialPosition': '連結財政状態計算書',
                'ConsolidatedStatementOfProfitOrLoss': '連結損益計算書',
                'ConsolidatedStatementOfFinancialPositionIFRS': '連結財政状態計算書',
                'ConsolidatedStatementOfProfitOrLossIFRS': '連結損益計算書',
                'ConsolidatedStatementOfCashFlowsIFRS': '連結キャッシュ・フロー計算書',
                'ConsolidatedStatementOfChangesInEquityIFRS': '連結株主資本等変動計算書',
                'ConsolidatedStatementOfComprehensiveIncomeIFRS': '連結包括利益計算書',
                'InvestmentsAndOtherAssets': '投資その他の資産',
                'Assets': '資産合計',
                'CurrentLiabilities': '流動負債',
                'NoncurrentLiabilities': '固定負債',
                'Liabilities': '負債合計',
                'NetAssets': '純資産合計',
                'LiabilitiesAndNetAssets': '負債純資産合計',
                'NetSales': '売上高',
                'Revenue': '売上収益',
                'NetSalesIFRS': '売上収益', 
                'RevenueIFRS': '売上収益',
                'CostOfSales': '売上原価',
                'GrossProfit': '売上総利益',
                'SellingGeneralAndAdministrativeExpenses': '販売費及び一般管理費',
                'SellingGeneralAndAdministrativeExpense': '販売費及び一般管理費',
                'OtherOperatingIncome': 'その他の営業収益',
                'OtherOperatingExpenses': 'その他の営業費用',
                'OtherOperatingExpense': 'その他の営業費用',
                # 【追加】IFRSの「その他の収益」「その他の費用」
                'OtherIncomeIFRS': 'その他の収益', 
                'OtherExpensesIFRS': 'その他の費用',
                # ついでに営業収益/費用も念のため追加
                'OtherOperatingIncomeIFRS': 'その他の営業収益',
                'OtherOperatingExpensesIFRS': 'その他の営業費用',
                'ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod': '持分法による投資利益',
                'OperatingIncome': '営業利益',
                'OperatingProfit': '営業利益',
                'OrdinaryIncome': '経常利益',
                'FinanceIncome': '金融収益',
                'FinancialIncome': '金融収益',
                'FinanceCosts': '金融費用',
                'FinancialExpenses': '金融費用',
                'FinanceExpenses': '金融費用',
                'ProfitBeforeTax': '税引前利益',
                'IncomeTaxExpense': '法人所得税費用',
                'Profit': '当期利益',
                'NetIncome': '当期利益',
                'ProfitLossAttributableToAbstract': '当期利益の帰属',
                'ProfitAttributableToOwnersOfParent': '親会社の所有者',
                'ProfitLossAttributableToOwnersOfParent': '親会社の所有者',
                'ProfitAttributableToNoncontrollingInterests': '非支配持分',
                'ProfitLossAttributableToNoncontrollingInterests': '非支配持分',
                'BasicEarningsPerShare': '基本的１株当たり当期利益（円）',
                'BasicEarningsLossPerShare': '基本的１株当たり当期利益（円）',
                # Priority IFRS items
                'CostOfSalesIFRS': '売上原価',
                'GrossProfitIFRS': '売上総利益',
                'SellingGeneralAndAdministrativeExpensesIFRS': '販売費及び一般管理費',
                'ShareOfProfitLossOfInvestmentsAccountedForUsingEquityMethodIFRS': '持分法による投資利益',
                'OperatingProfitLossIFRS': '営業利益',
                'FinanceIncomeIFRS': '金融収益',
                'FinanceCostsIFRS': '金融費用',
                'ProfitLossBeforeTaxIFRS': '税引前利益',
                'IncomeTaxExpenseIFRS': '法人所得税費用',
                'ProfitLossIFRS': '当期利益',
                'ProfitLossAttributableToOwnersOfParentIFRS': '親会社の所有者',
                'ProfitLossAttributableToNonControllingInterestsIFRS': '非支配持分',
                'BasicEarningsPerShareIFRS': '基本的１株当たり当期利益（円）'
            }
            
            parts = el.split('_')
            base_name = parts[-1] if len(parts) > 1 else el
            
            if base_name in common_dict:
                label = common_dict[base_name]
            else:
                label = labels_map.get(el)
                if not label:
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
            
            row_data = [indent_prefix + label, el]
            
            has_data = False
            for col_key in sorted_role_cols:
                val = all_years_data[role][full_path].get(col_key, "")
                # Clean numeric values
                if val:
                    # Handle full-width characters and commas
                    import unicodedata
                    val_clean = unicodedata.normalize('NFKC', str(val)).replace(',', '').strip()
                    try:
                        # Only convert if it looks numeric (don't convert plain text names)
                        if val_clean and not any(c.isalpha() for c in val_clean):
                            val = float(val_clean)
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
                    cell.number_format = r'#,##0_ ;[Red]\-#,##0 '
                    
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

    # シートの並び替え
    def get_sheet_order(title):
        is_note = '注記' in title
        group = 0
        if not is_note:
            group = 1 if '連結' in title else 3
        else:
            # 単体の財務諸表に対応する注記名かどうかで判定
            non_consolidated_notes = ['注記_貸借対照表', '注記_損益計算書', '注記_株主資本等変動計算書', '注記_包括利益計算書', '注記_キャッシュ・フロー計算書', '注記_製造原価明細書']
            if '連結' not in title and any(n in title for n in non_consolidated_notes):
                group = 4
            else:
                group = 2
                
        # 財務諸表の種別による並び順
        stmt_order = 99
        if '貸借対照表' in title or '財政状態' in title:
            stmt_order = 1
        elif '損益' in title or ('利益' in title and '包括' not in title):
            stmt_order = 2
        elif '包括' in title:
            stmt_order = 3
        elif '変動' in title:
            stmt_order = 4
        elif 'キャッシュ' in title:
            stmt_order = 5
            
        # 基準ごとの並び順 (日本基準が先)
        std_order = 2 if '(IFRS)' in title else 1
        
        return (group, stmt_order, std_order)
                
    wb._sheets.sort(key=lambda s: get_sheet_order(s.title))

    out_file = f'XBRL_横展開_{company_name}.xlsx'
    if output_dir:
        out_file = os.path.join(output_dir, out_file)
    wb.save(out_file)
    return out_file

def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_xbrl_to_excel.py <path_to_zip1> [<path_to_zip2> ...]", file=sys.stderr)
        sys.exit(1)
    process_xbrl_zips(sys.argv[1:])

if __name__ == "__main__":
    main()