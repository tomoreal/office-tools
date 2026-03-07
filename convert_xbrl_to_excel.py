import os
import sys
import zipfile
import tempfile
import shutil
import glob
import time
import json
import urllib.request
import re
from bs4 import BeautifulSoup
try:
    from lxml import etree
    HAS_LXML = True
except ImportError:
    import xml.etree.ElementTree as etree
    HAS_LXML = False

import gc

# Base directory for the script and caching
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Delay loading heavy libraries until needed (helps CGI performance)
HAS_BS4 = True # bs4 is at the top but BeautifulSoup is imported there.
HAS_PANDAS = False
HAS_OPENPYXL = False

# Control verbose logging via environment variable (default: enabled for debugging)
VERBOSE_LOGGING = os.environ.get('XBRL_VERBOSE', '1') == '1'

def debug_log(message):
    """Write message to a persistent debug log file for user visibility."""
    log_file = os.path.join(SCRIPT_DIR, 'convert_xbrl_debug.log')
    try:
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp} {message}\n")
    except:
        pass
    # Also print to stderr for server log visibility
    print(message, file=sys.stderr)

def vprint(*args, **kwargs):
    """Verbose print - only prints if VERBOSE_LOGGING is enabled."""
    if VERBOSE_LOGGING:
       msg = " ".join(map(str, args))
       debug_log(f"[VERBOSE] {msg}")


def safe_xpath(tree_or_elem, query, namespaces=None):
    """Safe XPath helper that works with both lxml and standard xml.etree.ElementTree.
    Note: ElementTree supports only a subset of XPath.
    """
    if HAS_LXML:
        return tree_or_elem.xpath(query, namespaces=namespaces)
    else:
        # Fallback for standard ElementTree (basic namespace handling)
        # Note: ET uses {url}prefix syntax for namespaces in findall
        # We try to convert basic queries but complex XPath might skip.
        try:
            # If it's a Tree object, get root first
            root = tree_or_elem.getroot() if hasattr(tree_or_elem, 'getroot') else tree_or_elem
            if namespaces:
                # Basic conversion for link:loc -> {http://...}loc if query is simple
                # For now, we return empty or try simple findall if query starts with //
                if query.startswith('//'):
                    tag = query.split(':')[-1]
                    return root.findall(f'.//{{*}}{tag}')
            return root.findall(query) 
        except:
            return []

# Common dimension axis and member mapping for Japanese fallback
COMMON_DIMENSION_MAPPING = {
    'ConsolidatedOrNonConsolidatedAxis': '連結・非連結',
    'ConsolidatedMember': '連結',
    'NonConsolidatedMember': '単体',
    'OperatingSegmentsAxis': '事業セグメント',
    'OperatingSegmentsMember': '事業セグメント',
    'BusinessSegmentsAxis': '事業セグメント',
    'BusinessSegmentsMember': '事業セグメント',
    'ReportableSegmentsAxis': '報告セグメント',
    'ReportableSegmentsMember': '報告セグメント',
    'EntityInformationAxis': '提出者情報',
    # Business Segment Standard Members (2024 Taxonomy)
    'OperatingSegmentsNotIncludedInReportableSegmentsAndOtherRevenueGeneratingBusinessActivitiesMember': '報告セグメント以外の全てのセグメント',
    'TotalOfReportableSegmentsAndOthersMember': '報告セグメント及びその他の合計',
    'UnallocatedAmountsAndEliminationMember': '全社・消去',
    'ReconcilingItemsMember': '調整項目',
}

# IFRS account name mapping to match commercial tools
IFRS_LABEL_MAPPING = {
    'jpigp_cor_RevenueIFRS': '売上高',
    'jpigp_cor_Revenue': '売上高',
    'jpigp_cor_ProfitLossBeforeTaxIFRS': '税引前当期利益',
    'jpigp_cor_ProfitLossAttributableToOwnersOfParentIFRS': '親会社の所有者',
    'jpigp_cor_ProfitLossAttributableToNonControllingInterestsIFRS': '非支配持分',
    'jpigp_cor_AssetsIFRS': '資産合計',
    'jpigp_cor_EquityIFRS': '資本合計',
    'jpigp_cor_EquityAttributableToOwnersOfParentIFRS': '親会社の所有者に帰属する持分合計',
    'jpigp_cor_OperatingProfitIFRS': '営業利益',
    'jpigp_cor_OperatingRevenueIFRS': '営業収益',
    'jpigp_cor_CostOfSalesIFRS': '売上原価',
    'jpigp_cor_GrossProfitIFRS': '売上総利益',
    'jpigp_cor_SellingGeneralAndAdministrativeExpensesIFRS': '販売費及び一般管理費',
    'jpigp_cor_InventoriesCAIFRS': '棚卸資産',
    'jpigp_cor_PropertyPlantAndEquipmentIFRS': '有形固定資産',
    'jpigp_cor_IntangibleAssetsIFRS': '無形資産',
    'jpigp_cor_CurrentAssetsIFRS': '流動資産合計',
    'jpigp_cor_NonCurrentAssetsIFRS': '非流動資産合計',
    'jpigp_cor_LiabilitiesIFRS': '負債合計',
}

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

def get_standard_labels(year, cache_dir=None):
    """Returns (all_labels, label_priorities) for the given taxonomy year.
    Uses cached standard_labels.json if it exists.
    """
    if cache_dir is None:
        cache_dir = os.path.join(SCRIPT_DIR, 'edinet_taxonomies')
    
    start_time = time.time()
    tax_dir = os.path.join(cache_dir, str(year))
    labels_cache_file = os.path.join(tax_dir, 'standard_labels.json')
    
    debug_log(f"Checking taxonomy cache: {labels_cache_file}")
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
        vprint(f"Taxonomy for year {year} not found in our known URL map.")
        return {}, {}
        
    tax_dir = os.path.join(cache_dir, str(year))
    labels_cache_file = os.path.join(tax_dir, 'standard_labels.json')
    
    # Try to load from cache
    if os.path.exists(labels_cache_file):
        try:
            with open(labels_cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, dict) and 'labels' in data:
                    res = data['labels'], data.get('priorities', {})
                    debug_log(f"SUCCESS: Loaded taxonomy cache for {year} in {time.time() - start_time:.2f}s")
                    return res
                else:
                    # Legacy format compatibility
                    priorities = {k: 50 for k in data}
                    debug_log(f"SUCCESS: Loaded legacy taxonomy cache for {year} in {time.time() - start_time:.2f}s")
                    return data, priorities
        except Exception as e:
            debug_log(f"ERROR: Cache read error for {year}: {e}")
            
    if not os.path.exists(tax_dir):
        try:
            os.makedirs(tax_dir, exist_ok=True)
            debug_log(f"Created taxonomy directory: {tax_dir}")
        except Exception as e:
            debug_log(f"WARNING: Could not create tax_dir {tax_dir}, falling back to /tmp: {e}")
            tax_dir = os.path.join('/tmp', 'edinet_taxonomies', str(year))
            try:
                os.makedirs(tax_dir, exist_ok=True)
            except: pass
            labels_cache_file = os.path.join(tax_dir, 'standard_labels.json')

    if not os.path.exists(labels_cache_file):
        zip_path = os.path.join(tax_dir, 'taxonomy.zip')
        if not os.path.exists(os.path.join(tax_dir, 'taxonomy')): # rudimentary check for extracted data
            vprint(f"Downloading EDINET taxonomy for {year} (takes a moment)...")
            try:
                urllib.request.urlretrieve(urls[year], zip_path)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    # Robust extraction: manually decode filenames using CP932 (shift_jis)
                    # to avoid Mojibake on Linux/Unix systems that default to UTF-8
                    for info in zip_ref.infolist():
                        try:
                            # infolist().filename is often bytes or interpreted as CP437
                            # We re-encode and decode correctly
                            filename_raw = info.filename.encode('cp437')
                            filename = filename_raw.decode('cp932')
                        except:
                            filename = info.filename
                        
                        target_path = os.path.join(tax_dir, filename)
                        if info.is_dir():
                            os.makedirs(target_path, exist_ok=True)
                        else:
                            os.makedirs(os.path.dirname(target_path), exist_ok=True)
                            with zip_ref.open(info) as source, open(target_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                os.remove(zip_path)
                
                # Canonical normalization: Rename any Mojibake "タクソノミ" folder to "taxonomy"
                for entry in os.listdir(tax_dir):
                    full_p = os.path.join(tax_dir, entry)
                    if os.path.isdir(full_p) and ('â^' in entry or 'タクソノミ' in entry):
                        new_p = os.path.join(tax_dir, 'taxonomy')
                        if not os.path.exists(new_p):
                            os.rename(full_p, new_p)
                            vprint(f"Normalized taxonomy directory: {entry} -> taxonomy")
            except Exception as e:
                vprint(f"Failed to download/extract taxonomy for {year}: {e}")
                return {}, {}
            
    vprint(f"Parsing EDINET taxonomy labels for {year}... (First run only)")
    lab_files = sorted(glob.glob(os.path.join(tax_dir, '**', '*_lab.xml'), recursive=True))
    all_labels = {}
    label_priorities = {} # {element_name: priority}

    # Track which taxonomy types we're loading
    taxonomy_types = set()

    for lf in lab_files:
        if 'deprecated' in lf or 'dep' in lf or '-en.xml' in lf:
            continue

        # Extract taxonomy type (jpigp, jppfs, jpcrp, etc.)
        basename = os.path.basename(lf)
        if '_lab.xml' in basename:
            tax_type = basename.split('_')[0]
            taxonomy_types.add(tax_type)

        try:
            # Determine taxonomy type from filename for domain-specific weighting
            tax_type = os.path.basename(lf).split('_')[0]
            parsed_labels, parsed_priorities = parse_labels_file(lf)
            
            for el, text in parsed_labels.items():
                prio = parsed_priorities.get(el, 99)
                
                # Element-Prefix Awareness: Boost priority if the taxonomy file matches the element's prefix
                # e.g. jppfs label for jppfs element is better than general label
                el_prefix = el.split('_')[0] if '_' in el else ""
                if el_prefix and el_prefix == tax_type:
                    prio -= 0.5 # Slight boost for domain-exact match
                
                current_prio = label_priorities.get(el, 100)
                if (el not in all_labels) or (prio < current_prio) or (prio == current_prio and text < all_labels.get(el, "")):
                    all_labels[el] = text
                    label_priorities[el] = prio
        except:
            pass

    # Report which taxonomies were loaded
    if taxonomy_types:
        vprint(f"  Loaded taxonomies: {', '.join(sorted(taxonomy_types))}")

    ifrs_count = sum(1 for k in all_labels if 'IFRS' in k)
    vprint(f"  Total labels: {len(all_labels)}, IFRS labels: {ifrs_count}")
            
    if all_labels:
        try:
            with open(labels_cache_file, 'w', encoding='utf-8') as f:
                json.dump({'labels': all_labels, 'priorities': label_priorities}, f, ensure_ascii=False, indent=2)
            debug_log(f"SUCCESS: Saved taxonomy cache to {labels_cache_file} in {time.time() - start_time:.2f}s")
        except Exception as e:
            debug_log(f"WARNING: Could not cache labels to {labels_cache_file}: {e}")
            
    return all_labels, label_priorities


def parse_labels_file(lab_file):
    """Parse an XBRL label linkbase using lxml for robust namespace handling.
    Returns (labels, priorities) where labels is a dict mapping element names to text,
    and priorities maps them to their best priority score.
    """
    labels = {}
    priorities = {}
    try:
        # Parse XML with lxml if available (ElementTree.XMLParser doesn't support 'recover')
        if HAS_LXML:
            parser = etree.XMLParser(recover=True)
            tree = etree.parse(lab_file, parser)
        else:
            tree = etree.parse(lab_file)
    except Exception as e:
        # If parsing fails, return empty mappings
        vprint(f"Error parsing {lab_file}: {e}")
        return labels, priorities

    # Namespace map for XBRL linkbase
    ns = {
        "link": "http://www.xbrl.org/2003/linkbase",
        "xlink": "http://www.w3.org/1999/xlink",
        "xml": "http://www.w3.org/XML/1998/namespace"
    }

    # 1. Locate all <link:loc> elements to map label IDs to element QNames
    href_to_label_id = {}
    for loc in safe_xpath(tree, "//link:loc", namespaces=ns):
        href = loc.get("{http://www.w3.org/1999/xlink}href")
        label_id = loc.get("{http://www.w3.org/1999/xlink}label")
        if href and label_id:
            # Element name may be a QName like jppfs_cor:CashAndDeposits
            element_name = href.split('#')[-1].replace(':', '_')
            href_to_label_id[label_id] = element_name

    # 2. Filter arcs to only concept‑label relationships (collect ALL associated resource IDs)
    label_id_to_res_ids = {}
    arc_xpath = "//link:labelArc[@xlink:arcrole='http://www.xbrl.org/2003/arcrole/concept-label']"
    for arc in safe_xpath(tree, arc_xpath, namespaces=ns):
        from_id = arc.get("{http://www.w3.org/1999/xlink}from")
        to_id = arc.get("{http://www.w3.org/1999/xlink}to")
        if from_id and to_id:
            if from_id not in label_id_to_res_ids:
                label_id_to_res_ids[from_id] = []
            label_id_to_res_ids[from_id].append(to_id)

    # 3. Gather label resources (<link:label>) with Japanese language
    res_id_to_text = {}
    res_id_to_priority = {}
    
    # Role priority: verboseLabel is standard for EDINET CSV output
    # XBRL Label Roles and their associated priority (lower is better)
    role_priority = {
        "http://www.xbrl.org/2003/role/verboseLabel": 1,
        "http://disclosure.edinet-fsa.go.jp/jpcrp/alt/role/label": 2, # EDINET industry-specific alternate
        "http://www.xbrl.org/2003/role/label": 3,
        "http://disclosure.edinet-fsa.go.jp/jppfs/ele/role/label": 4, # Electric Power
        "http://disclosure.edinet-fsa.go.jp/jppfs/gas/role/label": 4, # Gas
        "http://disclosure.edinet-fsa.go.jp/jppfs/sec/role/label": 4, # Securities
        "http://disclosure.edinet-fsa.go.jp/jppfs/ins/role/label": 4, # Insurance
        "http://disclosure.edinet-fsa.go.jp/jppfs/bnk/role/label": 4, # Banking
        "http://www.xbrl.org/2003/role/terseLabel": 5,
        "http://www.xbrl.org/2003/role/totalLabel": 10,
        "http://disclosure.edinet-fsa.go.jp/jpcrp/alt/role/totalLabel": 11,
    }

    GENERIC_LABELS = ('合計', '小計', '計', 'total', 'sum', 'subtotal', '金額')

    for res in safe_xpath(tree, "//link:label", namespaces=ns):
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
        # Phase 1: Skip penalty if it's the high-priority verboseLabel (priority 1)
        if priority > 1 and any(g in text.lower() for g in GENERIC_LABELS):
            priority += 50
            
        if (res_id not in res_id_to_text) or (priority < res_id_to_priority.get(res_id, 100)) or (priority == res_id_to_priority.get(res_id, 100) and text < res_id_to_text[res_id]):
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
    vprint(f"Parsing presentation linkbase... {os.path.basename(pre_file)}")
    try:
        # Use lxml for robust namespace handling if available
        if HAS_LXML:
            parser = etree.XMLParser(recover=True)
            tree = etree.parse(pre_file, parser)
        else:
            tree = etree.parse(pre_file)
    except Exception as e:
        vprint(f"Error parsing presentation linkbase: {e}")
        return {}

    ns = {
        "link": "http://www.xbrl.org/2003/linkbase",
        "xlink": "http://www.w3.org/1999/xlink"
    }

    # 1. Group by role URI first
    role_to_content = {} # {role_uri: {'locs': {label: element}, 'arcs': [arc_dicts]}}
    
    links = safe_xpath(tree, "//link:presentationLink", namespaces=ns)
    for link in links:
        role_uri = link.get("{http://www.w3.org/1999/xlink}role")
        if not role_uri:
            continue
        
        if role_uri not in role_to_content:
            role_to_content[role_uri] = {'locs': {}, 'arcs': []}
            
        # Map locators in this link
        locs = safe_xpath(link, "link:loc", namespaces=ns)
        for loc in locs:
            href = loc.get("{http://www.w3.org/1999/xlink}href")
            label = loc.get("{http://www.w3.org/1999/xlink}label")
            if href and label:
                # Normalize element name: replace ':' with '_' to match facts and labels
                element_name = href.split('#')[-1].replace(':', '_')
                role_to_content[role_uri]['locs'][label] = element_name
                
        # Map arcs in this link
        arcs = safe_xpath(link, "link:presentationArc", namespaces=ns)
        for arc in arcs:
            from_id = arc.get("{http://www.w3.org/1999/xlink}from")
            to_id = arc.get("{http://www.w3.org/1999/xlink}to")
            order = arc.get("order")
            pref_label = arc.get("preferredLabel")
            if from_id and to_id:
                role_to_content[role_uri]['arcs'].append({
                    'from': from_id,
                    'to': to_id,
                    'order': float(order) if order else 0.0,
                    'preferredLabel': pref_label
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
                    'order': arc['order'],
                    'preferredLabel': arc.get('preferredLabel')
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
    vprint(f"Parsing XBRL contexts and units... {os.path.basename(xbrl_file)}")
    try:
        # Use lxml for robust namespace handling if available
        if HAS_LXML:
            parser = etree.XMLParser(recover=True)
            tree = etree.parse(xbrl_file, parser)
        else:
            tree = etree.parse(xbrl_file)
    except Exception as e:
        vprint(f"Error parsing XBRL instance: {e}")
        return {}, {}

    # Standard namespaces for XBRL instance and dimensions
    ns = {
        "xbrli": "http://www.xbrl.org/2003/instance",
        "xbrldi": "http://xbrl.org/2006/xbrldi"
    }

    contexts = {}
    
    # 1. Parse contexts
    for ctx in safe_xpath(tree, "//xbrli:context", namespaces=ns):
        ctx_id = ctx.get('id')
        if not ctx_id:
            continue
            
        members = safe_xpath(ctx, ".//xbrldi:explicitMember", namespaces=ns)
        dim_vals = []
        for m in members:
            # Handle QNames in member text (e.g., jppfs_cor:EnergySegmentMember)
            m_text = m.text or ""
            
            # --- Axis Name Resolution ---
            dim_qname = m.get("dimension")
            dim_val = dim_qname.split(':')[-1] if dim_qname else ''
            
            # Resolve Axis Label
            prefixes = ['jpcrp_cor_', 'jppfs_cor_', 'jpigp_cor_', 'jpdei_cor_', '']
            axis_label = None
            if dim_val in COMMON_DIMENSION_MAPPING:
                axis_label = COMMON_DIMENSION_MAPPING[dim_val]
            else:
                for p in prefixes:
                    if p + dim_val in labels_map:
                        axis_label = labels_map[p + dim_val].replace(' [軸]', '').replace(' [項目]', '').replace(' [区分]', '').strip()
                        break
                if not axis_label:
                    axis_label = convert_camel_case_to_title(dim_val.replace('Axis', '')) if dim_val else ''

            # --- Member Name Resolution ---
            member_val = m_text.split(':')[-1]
            label = None
            if member_val in COMMON_DIMENSION_MAPPING:
                label = COMMON_DIMENSION_MAPPING[member_val]
            else:
                # Try with standard prefixes first
                for p in prefixes:
                    if p + member_val in labels_map:
                        label = labels_map[p + member_val]
                        break
                
                # If not found, try searching for any element ending with this member name in labels_map
                # (to catch standard elements from any taxonomy namespace)
                if not label:
                    suffix = '_' + member_val
                    for k, v in labels_map.items():
                        if k.endswith(suffix):
                            label = v
                            break
                    
            # Fallback for company specific segment names found in _lab.xml 
            if label: label = label.replace(' [メンバー]', '').replace(' [要素]', '').replace(' [区分]', '').strip()
            if not label:
                suffix = '_' + member_val
                for k, v in labels_map.items():
                    if k.endswith(suffix):
                        label = v
                        break
            
            if label: label = label.replace(' [メンバー]', '').replace(' [要素]', '').replace(' [区分]', '').strip()
            if not label:
                if member_val.endswith('Member'):
                    label = convert_camel_case_to_title(member_val.replace('Member', ''))
                else:
                    label = member_val
            
            # Combine Axis and Member if useful
            skip_axes = ('報告セグメント', 'セグメント情報', '事業セグメント', '会計基準', '連結・単体', '連結・非連結', 
                         'ConsolidatedOrNonConsolidated', 'OperatingSegments', 'BusinessSegments', 'ReportableSegments')
            if axis_label and not any(sa in axis_label.replace(' ', '') for sa in skip_axes):
                dim_vals.append(f"{axis_label}：{label}")
            else:
                dim_vals.append(label)
                
        dim_str = "、".join(dim_vals) if dim_vals else "全体"
        # Clean up verbose XBRL labels
        dim_str = dim_str.replace(' [メンバー]', '').replace('、報告セグメント', '').replace('非連結又は個別', '単体').replace('非連結', '単体')
        if dim_str == 'NonConsolidated' or dim_str == 'Non Consolidated':
            dim_str = '単体'
        if dim_str == 'Consolidated':
            dim_str = '連結'
            
        period_elem = safe_xpath(ctx, "xbrli:period", namespaces=ns)
        if period_elem:
            period_elem = period_elem[0]
            instant = safe_xpath(period_elem, "xbrli:instant", namespaces=ns)
            end_date = safe_xpath(period_elem, "xbrli:endDate", namespaces=ns)
            
            p_val = None
            start_val = None
            if instant:
                p_val = instant[0].text
            elif end_date:
                p_val = end_date[0].text
                start_elem = safe_xpath(period_elem, "xbrli:startDate", namespaces=ns)
                if start_elem:
                    start_val = start_elem[0].text
                
            if p_val:
                contexts[ctx_id] = (p_val, dim_str, start_val)
                
    units = {}
    # 2. Parse units
    for unit in safe_xpath(tree, "//xbrli:unit", namespaces=ns):
        unit_id = unit.get('id')
        if not unit_id:
            continue
        
        is_jpy = False
        # Only consider simple units (non‑divide) for JPY amount identification
        if not safe_xpath(unit, "xbrli:divide", namespaces=ns):
            measure = safe_xpath(unit, ".//xbrli:measure", namespaces=ns)
            if measure and 'JPY' in (measure[0].text or ""):
                is_jpy = True
                
        units[unit_id] = is_jpy
                
    return contexts, units

def parse_ixbrl_facts(ixbrl_files, contexts, units):
    t_start = time.time()
    parser_info = 'lxml' if HAS_LXML else 'html.parser'
    debug_log(f"Starting Inline XBRL parsing using {parser_info} for {len(ixbrl_files)} files")
    facts = []
    
    for f in ixbrl_files:
        size_mb = os.path.getsize(f) / (1024 * 1024)
        debug_log(f"  Parsing {os.path.basename(f)} ({size_mb:.2f} MB)...")
        try:
            with open(f, 'r', encoding='utf-8', errors='replace') as file:
                content = file.read()
            
            if HAS_LXML:
                try:
                    from lxml import html
                    tree = html.fromstring(content)
                    # Use a more robust way to find tags that works with or without namespace awareness
                    tags = [t for t in tree.iter() if any(x in (t.tag if isinstance(t.tag, str) else "").lower() for x in ('nonfraction', 'nonnumeric'))]
                except Exception as e:
                    debug_log(f"  LXML fast-path failed: {e}. Falling back to BS4.")
                    HAS_LXML_LOCAL = False
                else:
                    HAS_LXML_LOCAL = True
            else:
                HAS_LXML_LOCAL = False

            if not HAS_LXML_LOCAL:
                soup = BeautifulSoup(content, 'html.parser')
                def is_ix_tag(tag):
                    if not tag.name: return False
                    local = tag.name.split(':')[-1].lower()
                    return local in ('nonfraction', 'nonnumeric')
                tags = soup.find_all(is_ix_tag)

            elem_order_in_file = 0
            for tag in tags:
                if HAS_LXML_LOCAL:
                    # Access attributes robustly (handle namespaced keys like {uri}name or ix:name)
                    # LXML uses {uri}attribute_name format for namespaced attributes
                    attrs = {}
                    for k, v in tag.attrib.items():
                        # Extract local name: handle {uri}name and prefix:name
                        local_k = k
                        if '}' in local_k:
                            local_k = local_k.split('}')[-1]
                        if ':' in local_k:
                            local_k = local_k.split(':')[-1]
                        attrs[local_k.lower()] = v
                    ctx_ref = attrs.get('contextref')
                    if not ctx_ref or ctx_ref not in contexts: continue
                    
                    element_name = attrs.get('name')
                    if not element_name: continue
                    if ':' in element_name:
                        element_name = element_name.replace(':', '_')
                    
                    value = tag.text_content().strip() if hasattr(tag, 'text_content') else (tag.text or "").strip()
                    local_name = tag.tag.split('}')[-1].lower() if isinstance(tag.tag, str) and '}' in tag.tag else (tag.tag.split(':')[-1].lower() if isinstance(tag.tag, str) else "")
                    
                    unit_ref = attrs.get('unitref')
                    scale = attrs.get('scale', '0')
                    sign = attrs.get('sign', '')
                else:
                    # BS4 path
                    ctx_ref = None
                    for k, v in tag.attrs.items():
                        if k.lower() == 'contextref':
                            ctx_ref = v
                            break
                    if not ctx_ref or ctx_ref not in contexts: continue
                    
                    element_name = None
                    for k, v in tag.attrs.items():
                        if k.lower() == 'name':
                            element_name = v
                            break
                    if not element_name: continue
                    if ':' in element_name:
                        element_name = element_name.replace(':', '_')
                    
                    value = tag.get_text().strip()
                    local_name = tag.name.split(':')[-1].lower()

                if local_name == 'nonnumeric':
                    # Skip massive text blocks only if it's explicitly a TextBlock element
                    if 'TextBlock' in element_name:
                        continue
                
                valStr = ""
                if local_name == 'nonfraction':
                    if not HAS_LXML_LOCAL:
                        unit_ref = None
                        for k, v in tag.attrs.items():
                            if k.lower() == 'unitref':
                                unit_ref = v
                                break
                        scale = '0'
                        sign = ''
                        for k, v in tag.attrs.items():
                            if k.lower() == 'scale': scale = v
                            if k.lower() == 'sign': sign = v

                    is_jpy = units.get(unit_ref, False) if unit_ref else False
                    clean_val = value.replace(',', '').replace('△', '-').replace('▲', '-').replace('(', '-').replace(')', '').strip()
                    
                    try:
                        amt = float(clean_val)
                        if sign == '-': amt *= -1
                        amt *= (10 ** int(scale or 0))
                        if is_jpy: amt /= 1000000.0
                        valStr = str(int(amt)) if amt.is_integer() else str(amt)
                    except:
                        valStr = value
                else:
                    valStr = value
                    
                f_data = {
                    'element': element_name,
                    'context': ctx_ref,
                    'period': contexts[ctx_ref][0],
                    'dimension': contexts[ctx_ref][1],
                    'value': valStr,
                    'source_file': f,
                    'elem_order': elem_order_in_file
                }
                if contexts[ctx_ref][2]: # start_date
                    f_data['start_date'] = contexts[ctx_ref][2]
                facts.append(f_data)
                elem_order_in_file += 1
            
            if elem_order_in_file > 0:
                vprint(f"  Extracted {elem_order_in_file} facts from {os.path.basename(f)}")
            
            if not HAS_LXML_LOCAL:
                soup.decompose()
                del soup
            else:
                del tree
            gc.collect()

        except Exception as e:
            debug_log(f"ERROR: Error parsing file {f}: {e}")
                
    debug_log(f"COMPLETED: Parsed all Inline XBRL facts in {time.time() - t_start:.2f}s")
    return facts




def create_hierarchy(parent_child_arcs):
    """Create a flattened list representing the hierarchy traversal."""
    # Group by parent to easily find children
    adj = {}
    for arc in parent_child_arcs:
        p = arc['parent']
        if p not in adj: adj[p] = []
        adj[p].append(arc)
    
    # Sort children by order, with stable tie-breaking on child name
    for p in adj:
        adj[p].sort(key=lambda x: (x['order'], x['child']))
        
    roots = set(arc['parent'] for arc in parent_child_arcs)
    children = set(arc['child'] for arc in parent_child_arcs)
    top_roots = sorted(list(roots - children))
    
    if not top_roots and parent_child_arcs:
        # If circular or no clear root, pick the parent of the first arc
        top_roots = [parent_child_arcs[0]['parent']]
        
    ordered_items = []
    seen = set()
    
    def traverse(node_name, path, depth, pref_label=None):
        # Use a tuple of (node_name, pref_label) to allow the same element 
        # to appear multiple times if it has different preferred labels 
        # (common in Cash Flow for beginning/ending balance)
        node_id = (node_name, pref_label, depth)
        if node_id in seen: return
        seen.add(node_id)
        
        full_path = path + "::" + node_name
        if pref_label:
            full_path += f"|{pref_label}"
            
        ordered_items.append((node_name, full_path, depth, pref_label))
        
        if node_name in adj:
            for arc in adj[node_name]:
                traverse(arc['child'], full_path, depth + 1, arc.get('preferredLabel'))
                
    for root in top_roots:
        traverse(root, "", 0)
        
    return ordered_items

def process_xbrl_zips(zip_paths, output_dir=None):
    overall_start = time.time()
    if not zip_paths:
        return None
    zip_paths = sorted(zip_paths)
        
    global HAS_PANDAS, HAS_OPENPYXL
    # Delay loading heavy libraries inside the function
    try:
        import pandas as pd
        HAS_PANDAS = True
    except ImportError:
        pd = None
        HAS_PANDAS = False

    try:
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        HAS_OPENPYXL = True
    except ImportError:
        Workbook = None
        HAS_OPENPYXL = False

    if not HAS_OPENPYXL:
        print("Error: openpyxl is not installed. Excel generation is impossible.", file=sys.stderr)
        return None
        
    global_element_period_values = {} # {element: {col_key: value}}
    merged_trees = {} # {role_name: {(parent, child): order}}
    seen_children_in_role = {} # {role_name: set(children)}
    labels_map = {} # {element: label_text}
    labels_map_priorities = {} # {element: priority}
    
    periods_seen = set()
    all_facts = []  # Accumulate facts across all zips for fallback logic

    # Use provided output_dir for temp files if possible to avoid permission issues in /tmp
    parent_temp_dir = output_dir if output_dir and os.path.exists(output_dir) else None
    temp_base = tempfile.mkdtemp(dir=parent_temp_dir)
    
    try:
        # Phase 3.5: Parallel processing of ZIP files
        from concurrent.futures import ThreadPoolExecutor
        
        def process_single_zip(zip_idx, zip_path):
            thread_labels = {}
            thread_priorities = {}
            thread_facts = []
            thread_periods = set()
            thread_values = {} # {el: {col: val}}
            
            debug_log(f"Starting worker for {os.path.basename(zip_path)}")
            if not os.path.exists(zip_path):
                return None
                
            extract_dir = os.path.join(temp_base, f"zip_{zip_idx}")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                
            subdirs = [d for d in os.listdir(extract_dir) if os.path.isdir(os.path.join(extract_dir, d))]
            if len(subdirs) == 1:
                extract_dir = os.path.join(extract_dir, subdirs[0])
                
            xbrl_files = find_xbrl_files(extract_dir)
            if not xbrl_files:
                return None
                
            taxonomy_year = None
            if xbrl_files['pre']:
                with open(xbrl_files['pre'], 'r', encoding='utf-8') as f:
                    content = f.read(4000)
                    m = re.search(r'http://disclosure\.edinet-fsa\.go\.jp/taxonomy/[a-z]+(?:_[a-z]+)?/(\d{4})-\d{2}-\d{2}', content)
                    if m:
                        year_str = m.group(1)
                        taxonomy_year = '2021' if year_str == '2020' else year_str
            
            if taxonomy_year:
                std_labels, std_priorities = get_standard_labels(taxonomy_year)
                thread_labels.update(std_labels)
                thread_priorities.update(std_priorities)

            for lf in xbrl_files.get('lab', []):
                local_labels, local_priorities = parse_labels_file(lf)
                for k, v in local_labels.items():
                    p = local_priorities.get(k, 99) - 1
                    if k not in thread_labels or p < thread_priorities.get(k, 100):
                        thread_labels[k] = v
                        thread_priorities[k] = p
            
            # Phase 2: Demote IFRS mapping priority
            for el_name, alias in IFRS_LABEL_MAPPING.items():
                if el_name not in thread_labels or 20 < thread_priorities.get(el_name, 100):
                    thread_labels[el_name] = alias
                    thread_priorities[el_name] = 20
            
            contexts, units = parse_instance_contexts_and_units(xbrl_files['xbrl'], thread_labels)
            
            # Phase 3: Selective Parsing (DISABLED for completeness)
            all_ix_files = glob.glob(os.path.join(extract_dir, 'XBRL', 'PublicDoc', '*_ixbrl.htm'))
            ix_files = sorted(all_ix_files)

            facts = parse_ixbrl_facts(ix_files, contexts, units) # Corrected: pass units, not labels
            thread_facts.extend(facts)
            debug_log(f"Worker for {os.path.basename(zip_path)} found {len(facts)} facts in {len(ix_files)} files")
            
            for f in facts:
                el = f['element']
                period = f['period']
                dim = f.get('dimension', '')
                val = f['value']
                dim_label = dim if dim else "全体"
                col_key = (dim_label, period)
                if el not in thread_values: thread_values[el] = {}
                thread_values[el][col_key] = val
                thread_periods.add(col_key)
                # Store extra metadata (startDate) for periodStartLabel lookup
                if 'start_date' in f:
                    if '_metadata' not in thread_values: thread_values['_metadata'] = {}
                    thread_values['_metadata'][col_key] = f['start_date']
                
            trees = parse_presentation_linkbase(xbrl_files['pre'])
            
            return {
                'labels': thread_labels,
                'priorities': thread_priorities,
                'facts': thread_facts,
                'periods': thread_periods,
                'values': thread_values,
                'trees': trees
            }

        # Multi-threading for performance (I/O and C-based lxml parsing)
        # Use a maximum of 4 workers to avoid memory exhaustion in CGI
        with ThreadPoolExecutor(max_workers=min(len(zip_paths), 4)) as executor:
            def process_single_zip_wrapper(p):
                try:
                    return process_single_zip(p[0], p[1])
                except Exception as e:
                    debug_log(f"Worker failed for {p[1]}: {e}")
                    return None
            results = list(executor.map(process_single_zip_wrapper, enumerate(zip_paths)))
            
        for res in results:
            if not res: continue
            # Merge labels with priorities
            for k, v in res['labels'].items():
                p = res['priorities'].get(k, 100)
                if k not in labels_map or p < labels_map_priorities.get(k, 101):
                    labels_map[k] = v
                    labels_map_priorities[k] = p
            
            # Merge facts, periods, and values
            all_facts.extend(res['facts'])
            periods_seen.update(res['periods'])
            for el, vals in res['values'].items():
                if el not in global_element_period_values:
                    global_element_period_values[el] = {}
                for col_key, new_val in vals.items():
                    old_val = global_element_period_values[el].get(col_key)
                    if old_val is None:
                        global_element_period_values[el][col_key] = new_val
                    else:
                        # Tie-breaking: prefer values that look more precise (more decimals)
                        # This happens when the same fact appears in a table (precise) and a note (rounded)
                        # Or just be deterministic based on zip file order (already sorted)
                        def get_precision(s):
                            if not s or '.' not in s: return 0
                            return len(s.split('.')[-1])
                        if get_precision(new_val) > get_precision(old_val):
                            global_element_period_values[el][col_key] = new_val
            
            # Merge presentation trees
            for role, tree_arcs in res['trees'].items():
                base_name = role.split('_')[-1]
                # Merge SegmentInformation variants into a single role
                sub_role_idx = 0
                if 'SegmentInformation' in base_name and '-' in base_name:
                    parts = base_name.rsplit('-', 1)
                    if parts[1].isdigit():
                        sub_role_idx = int(parts[1]) * 1000
                    role = parts[0]
                
                # Filter relevant roles
                if not (base_name.startswith('Consolidated') or base_name.startswith('Statement') or 
                        base_name.startswith('BalanceSheet') or base_name.startswith('Notes') or 
                        'BusinessResults' in base_name):
                    continue
                
                if role not in merged_trees:
                    merged_trees[role] = {}
                    seen_children_in_role[role] = set()
                    
                for arc in tree_arcs:
                    p, c, o, pl = arc['parent'], arc['child'], arc['order'], arc.get('preferredLabel')
                    # Unique key including preferredLabel to allow duplicates in CF statements
                    arc_key = (p, c, pl)
                    merged_trees[role][arc_key] = float(o) + sub_role_idx
        
        debug_log(f"Merged total: {len(all_facts)} facts, {len(periods_seen)} periods, {len(merged_trees)} tree roles")
    except Exception as e:
        debug_log(f"ERROR: Overall processing failure: {e}")
        import traceback
        debug_log(traceback.format_exc())
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
            
            # For IFRS combined filings (typically 0105010), split into separate statements
            if doc_code == '0105010':
                headings_to_roles = {
                    'ConsolidatedStatementOfFinancialPositionIFRSHeading': 'rol_ConsolidatedStatementOfFinancialPositionIFRS',
                    'ConsolidatedStatementOfProfitOrLossIFRSHeading': 'rol_ConsolidatedStatementOfProfitOrLossIFRS',
                    'ConsolidatedStatementOfProfitOrLossAndOtherComprehensiveIncomeIFRSHeading': 'rol_ConsolidatedStatementOfProfitOrLossIFRS',
                    'ConsolidatedStatementOfCashFlowsIFRSHeading': 'rol_ConsolidatedStatementOfCashFlowsIFRS',
                    'ConsolidatedStatementOfChangesInEquityIFRSHeading': 'rol_ConsolidatedStatementOfChangesInEquityIFRS',
                    'StatementOfFinancialPositionIFRSHeading': 'rol_ConsolidatedStatementOfFinancialPositionIFRS',
                    'StatementOfProfitOrLossIFRSHeading': 'rol_ConsolidatedStatementOfProfitOrLossIFRS',
                    # Backup markers
                    'RevenueIFRS': 'rol_ConsolidatedStatementOfProfitOrLossIFRS',
                    'NetSalesIFRS': 'rol_ConsolidatedStatementOfProfitOrLossIFRS',
                    'NetCashProvidedByUsedInOperatingActivitiesIFRS': 'rol_ConsolidatedStatementOfCashFlowsIFRS',
                    'RetainedEarningsIFRS': 'rol_ConsolidatedStatementOfChangesInEquityIFRS',
                }
                
                curr_role = 'rol_ConsolidatedStatementOfFinancialPositionIFRS'
                curr_arcs = []
                roles_created = set()
                
                for elem, _order in sorted_elems:
                    base = elem.split('_')[-1]
                    if base in headings_to_roles:
                        new_role = headings_to_roles[base]
                        if new_role != curr_role:
                            if curr_arcs:
                                merged_trees[curr_role] = {(a['parent'], a['child'], a['preferredLabel']): a['order'] for a in curr_arcs}
                                print(f"[Fallback-Split] Created synthetic role {curr_role} (Phase 1)", file=sys.stderr)
                                roles_created.add(curr_role)
                            curr_role = new_role
                            curr_arcs = []
                    curr_arcs.append({'parent': curr_role, 'child': elem, 'order': float(_order), 'preferredLabel': None})
                if curr_arcs:
                    merged_trees[curr_role] = {(a['parent'], a['child'], a['preferredLabel']): a['order'] for a in curr_arcs}
                    print(f"[Fallback-Split] Created synthetic role {curr_role} (Phase 1)", file=sys.stderr)
            else:
                virtual_root = role_name
                arcs = []
                for elem, _order in sorted_elems:
                    arcs.append({'parent': virtual_root, 'child': elem, 'order': float(_order), 'preferredLabel': None})
                
                if arcs:
                    merged_trees[role_name] = {(a['parent'], a['child'], a['preferredLabel']): a['order'] for a in arcs}
                    print(f"[Fallback] Created synthetic role {role_name} from {doc_code} (Phase 1)", file=sys.stderr)

    all_years_data = {} # {role_name: {hierarchical_key: {period: value}}}
    role_to_order = {} # {role_name: [hierarchical_key1, ...]}
    
    for role, pd_dict in merged_trees.items():
        tree_arcs = [{'parent': p, 'child': c, 'order': o, 'preferredLabel': pl} for (p, c, pl), o in pd_dict.items()]
        ordered_items = create_hierarchy(tree_arcs)
        
        all_years_data[role] = {}
        role_to_order[role] = []
        for el, full_path, depth, pref_label in ordered_items:
            role_to_order[role].append((full_path, pref_label))
            all_years_data[role][full_path] = {}
            if el in global_element_period_values:
                for period, val in global_element_period_values[el].items():
                    all_years_data[role][full_path][period] = val

    # Try to find company name for filename
    company_name = "企業名不明"
    name_suffixes = ['CompanyNameCoverPage', 'EntityNameCompanyName', 'EntityNameEntityReportingName']
    for suffix in name_suffixes:
        found = False
        for el_name, vals in global_element_period_values.items():
            if el_name.endswith(suffix):
                if vals:
                    # Pick a value (latest period if possible)
                    sorted_keys = sorted(vals.keys(), key=lambda x: x[1] if isinstance(x, tuple) else x, reverse=True)
                    company_name = vals[sorted_keys[0]]
                    found = True
                    break
        if found: break
    
    # Debug: log top elements to see what was captured
    if VERBOSE_LOGGING or company_name == "企業名不明":
        top_el = sorted(list(global_element_period_values.keys()))[:30]
        debug_log(f"DEBUG: Company discovery failed. Top 30 elements: {top_el}")

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
            'BusinessResultsOfReportingCompany': '主要な経営指標等の推移（単体）',
            # IFRS Notes
            'InventoriesConsolidatedFinancialStatementsIFRS': '棚卸資産',
            'PropertyPlantAndEquipmentConsolidatedFinancialStatementsIFRS': '有形固定資産',
            'GoodwillAndIntangibleAssetsConsolidatedFinancialStatementsIFRS': 'のれん及び無形資産',
            'SellingGeneralAndAdministrativeExpensesConsolidatedFinancialStatementsIFRS': '販売費及び一般管理費',
            'FinanceIncomeAndFinanceCostsConsolidatedFinancialStatementsIFRS': '金融収益及び金融費用',
            'TradeAndOtherReceivablesConsolidatedFinancialStatementsIFRS': '営業債権及びその他の債権',
            'TradeAndOtherPayablesConsolidatedFinancialStatementsIFRS': '営業債務及びその他の債務',
            'OtherInvestmentsConsolidatedFinancialStatementsIFRS': 'その他の投資',
            'ExpensesByNatureConsolidatedFinancialStatementsIFRS': '費用の性質別内訳'
        }
        
        japanese_name = sheet_mapping.get(base_name)
        if not japanese_name:
            if base_name.startswith('Notes'):
                sub_name = base_name[5:] # remove 'Notes'
                # Dynamic lookup in labels_map for element names based on base_name
                # Try multiple prefixes for IFRS and J-GAAP
                prefixes = ["jpigp_cor_", "jpcrp_cor_", "jppfs_cor_"]
                # Possible variations: Prefix + role_name + suffix, or Prefix + sub_name + suffix
                search_terms = []
                for p in prefixes:
                    for suffix in ["Heading", "TextBlock", ""]:
                        search_terms.append(f"{p}{base_name}{suffix}")
                        if base_name.startswith('Notes'):
                            search_terms.append(f"{p}{base_name[5:]}{suffix}")
                        else:
                            search_terms.append(f"{p}Notes{base_name}{suffix}")
                
                for el in search_terms:
                    if el in labels_map:
                        raw_label = labels_map[el]
                        # Clean up: remove prefixes and standardize
                        # Example: "注記事項－..." or suffix phrases
                        clean_label = raw_label.split('、')[0].split(' [')[0].replace('注記事項－', '').strip()
                        if clean_label:
                            japanese_name = '注記_' + clean_label
                            break

                if not japanese_name:
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
        for full_path_data in ordered_keys:
            full_path, pref_label = full_path_data
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
        seen_labels = set() # Track labels for deduplication in each worksheet
        
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
        
        for full_path_data in ordered_keys:
            full_path, pref_label = full_path_data
            # Extract element name to get label
            el = full_path.split('::')[-1]
            if '|' in el: el = el.split('|')[0]
            
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
                'OtherIncomeIFRS': 'その他の収益', 
                'OtherExpensesIFRS': 'その他の費用',
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
                'ProfitLossBeforeTaxIFRS': '（税引前当期損益）',
                'IncomeTaxExpenseIFRS': '法人所得税費用',
                'ProfitLossIFRS': '当期利益',
                'ProfitLossAttributableToOwnersOfParentIFRS': '親会社の所有者',
                'ProfitLossAttributableToNonControllingInterestsIFRS': '非支配持分',
                'BasicEarningsPerShareIFRS': '基本的１株当たり当期利益（円）'
            }
            
            parts = el.split('_')
            base_name = parts[-1] if len(parts) > 1 else el
            
            segment_dict = {
                # IFRS Specific Overrides for Segments
                'NetSalesIFRS': '売上高', 
                'IntersegmentSalesIFRS': 'セグメント間売上高',
                'ProfitLossBeforeTaxIFRS': '（税引前当期損益）',
                'AssetsIFRS': 'セグメント資産',
                'DepreciationAndAmortizationOperatingExpensesIFRS': '減価償却費及び償却費',
                'ImpairmentLossesOnNonFinancialAssetsPLIFRS': '非金融資産の減損損失',
                'OtherIncomeAndExpensesNetIFRS': 'その他の損益',
                'ShareOfProfitLossOfInvestmentsAccountedForUsingEquityMethodIFRS': '持分法による投資損益',
                'CapitalExpendituresIFRS': '資本的支出',
                'InvestmentsAccountedForUsingEquityMethodIFRS': '持分法で会計処理されている投資',
                'FinanceIncomeIFRS': '金融収益',
                'FinanceCostsIFRS': '金融費用',
                'ExternalRevenueIFRS': '外部売上高',
                'IntersegmentRevenueIFRS': 'セグメント間売上高',
                'SegmentProfitLossIFRS': 'セグメント利益又は損失（△）',
                'SegmentAssetsIFRS': 'セグメント資産',
                'OtherInformationIFRS': 'その他の情報',
                'DepreciationAndAmortisationIFRS': '減価償却費及び償却費',
                'OtherProfitLossIFRS': 'その他の損益',
                
                # J-GAAP Specific Overrides for Segments
                'NetSales': '計',
                'OrdinaryIncome': 'セグメント利益又は損失（△）',
                'OperatingIncome': 'セグメント利益又は損失（△）',
                'Assets': 'セグメント資産',
                'Liabilities': 'セグメント負債',
                'SalesToExternalCustomers': '外部顧客への売上高',
                'IntersegmentSalesOrTransfers': 'セグメント間の内部売上高又は振替高',
                'SegmentProfitLoss': 'セグメント利益又は損失（△）',
                'SegmentAssets': 'セグメント資産',
                'SegmentLiabilities': 'セグメント負債',
                'OtherItems': 'その他の項目',
                'DepreciationAndAmortization': '減価償却費',
                'AmortizationOfGoodwill': 'のれんの償却額',
                'InterestIncome': '受取利息',
                'InterestExpenses': '支払利息',
                'ShareOfProfitLossOfEntitiesAccountedForUsingEquityMethod': '持分法投資利益又は損失（△）',
                'InvestmentsInEntitiesAccountedForUsingEquityMethod': '持分法適用会社への投資額',
                'IncreaseInPropertyPlantAndEquipmentAndIntangibleAssets': '有形固定資産及び無形固定資産の増加額'
            }
            
            if is_segment and base_name in segment_dict:
                label = segment_dict[base_name]
            elif base_name in common_dict:
                label = common_dict[base_name]
            else:
                label = labels_map.get(el)
                if label: label = label.replace(' [メンバー]', '').replace(' [要素]', '').replace(' [区分]', '').strip()
            if not label:
                    label = convert_camel_case_to_title(base_name)
            
            # Append suffix for Cash Flow balances
            is_cf_sheet = 'キャッシュ・フロー' in sheet_name
            if is_cf_sheet:
                if pref_label == 'http://www.xbrl.org/2003/role/periodStartLabel':
                    label += "（期首残高）"
                elif pref_label in ('http://www.xbrl.org/2003/role/periodEndLabel', 'http://www.xbrl.org/2003/role/totalLabel'):
                    # Only append (期末残高) if the element is likely a balance element
                    if 'CashAndCashEquivalents' in el:
                        label += "（期末残高）"
                    
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
            
            # --- セグメント情報の場合の重複排除 ---
            if is_segment:
                if label in seen_labels:
                    continue
                seen_labels.add(label)
                    
            # Indent based on depth
            depth = len(full_path.split('::')) - 1
            indent_prefix = "　" * depth
            
            row_data = [indent_prefix + label, el]
            
            has_numeric_data = False
            has_data = False
            for col_key in sorted_role_cols:
                # SPECIAL HANDLING for Cash Flow Beginning Balance
                # If this row is a periodStartLabel in a Cash Flow sheet, we pull data from the prior period's instant.
                val = ""
                is_cf_sheet = 'キャッシュ・フロー' in sheet_name
                is_start_row = pref_label == 'http://www.xbrl.org/2003/role/periodStartLabel'
                
                if is_cf_sheet and is_start_row:
                    # Find the startDate of the current column
                    current_start_date = global_element_period_values.get('_metadata', {}).get(col_key)
                    if current_start_date:
                        # Find the corresponding instant
                        # Cases:
                        # 1. Instant is exactly at current_start_date (e.g. 2024-04-01)
                        # 2. Instant is at the end of the previous day (e.g. 2024-03-31)
                        dates_to_try = [current_start_date]
                        try:
                            from datetime import datetime, timedelta
                            dt = datetime.strptime(current_start_date, '%Y-%m-%d')
                            prev_day = (dt - timedelta(days=1)).strftime('%Y-%m-%d')
                            dates_to_try.append(prev_day)
                        except:
                            pass
                        
                        found_val = False
                        for t_date in dates_to_try:
                            target_col_key = (col_key[0], t_date)
                            if el in global_element_period_values and target_col_key in global_element_period_values[el]:
                                val = global_element_period_values[el][target_col_key]
                                found_val = True
                                break
                        
                        if not found_val:
                            # Search for any dimension matching if exact dim fails
                            for t_date in dates_to_try:
                                for k, v in global_element_period_values.get(el, {}).items():
                                    if k[1] == t_date: # Match date, ignore dimension if needed? 
                                        # (Actually dimension should match, but sometimes it's "全体" vs "連結")
                                        val = v
                                        found_val = True
                                        break
                                if found_val: break
                
                # Only use fallback if this isn't a CF start row (to avoid pulling ending balance)
                if val == "" and not (is_cf_sheet and is_start_row):
                    val = all_years_data[role][full_path].get(col_key, "")
                
                # Clean numeric values
                if val:
                    # Handle full-width characters and commas
                    import unicodedata
                    val_clean = unicodedata.normalize('NFKC', str(val)).replace(',', '').strip()
                    try:
                        if val_clean and not any(c.isalpha() for c in val_clean):
                            val = float(val_clean)
                            has_numeric_data = True
                    except:
                        pass
                row_data.append(val)
                if val != "":
                    has_data = True
                    
            if has_data: # Only append rows that have at least one value across columns
                # --- セグメント情報の場合の文字情報の除外 ---
                if is_segment:
                    if not has_numeric_data:
                        continue
                ws.append(row_data)

        # Apply formatting and column widths
        ratio_elements = {
            # Standard & Japanese GAAP
            'EquityToAssetRatioSummaryOfBusinessResults',
            'RateOfReturnOnEquitySummaryOfBusinessResults',
            'CapitalAdequacyRatioInternationalStandardSummaryOfBusinessResults',
            'CapitalAdequacyRatioDomesticStandardSummaryOfBusinessResults',
            'CapitalAdequacyRatioBISStandardSummaryOfBusinessResults',
            'CapitalAdequacyRatioDomesticStandard2SummaryOfBusinessResults',
            'PayoutRatioSummaryOfBusinessResults',
            
            # IFRS Variations
            'RatioOfOwnersEquityToGrossAssetsIFRSSummaryOfBusinessResults',
            'RateOfReturnOnEquityIFRSSummaryOfBusinessResults',
            
            # JMIS Variations
            'RatioOfOwnersEquityToGrossAssetsJMISSummaryOfBusinessResults',
            'RateOfReturnOnEquityJMISSummaryOfBusinessResults',
            
            # US GAAP Variations
            'EquityToAssetRatioUSGAAPSummaryOfBusinessResults',
            'RateOfReturnOnEquityUSGAAPSummaryOfBusinessResults',
            
            # Industry Specific (Insurance, etc.)
            'NetLossRatioSummaryOfBusinessResultsINS',
            'NetOperatingExpenseRatioSummaryOfBusinessResultsINS'
        }

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            # Element name is in column B (index 1)
            el_name = row[1].value if len(row) > 1 else None
            is_ratio = False
            if el_name:
                el_str = str(el_name)
                # Handle prefixes: Namespace:Element or Namespace_Element
                # We check if the element name ends with our target ratio element name,
                # ensuring it is either a perfect match or preceded by a separator.
                for r in ratio_elements:
                    if el_str == r or el_str.endswith(':' + r) or el_str.endswith('_' + r):
                        is_ratio = True
                        break

            for cell in row:
                if isinstance(cell.value, (int, float)):
                    if is_ratio:
                        cell.number_format = '0.0%'
                    else:
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
    t_save = time.time()
    wb.save(out_file)
    debug_log(f"SUCCESS: Excel saved to {out_file} in {time.time() - t_save:.2f}s")
    debug_log(f"TOTAL: process_xbrl_zips completed in {time.time() - overall_start:.2f}s")
    return out_file

def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_xbrl_to_excel.py <path_to_zip1> [<path_to_zip2> ...]", file=sys.stderr)
        sys.exit(1)
    process_xbrl_zips(sys.argv[1:])

if __name__ == "__main__":
    main()