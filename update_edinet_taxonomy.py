#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EDINET Taxonomy Dictionary Auto-Update Script

This script automatically downloads the latest EDINET taxonomy from FSA
and regenerates the edinet_taxonomy_dict.py file.

Usage:
    python3 update_edinet_taxonomy.py [--force]

Options:
    --force    Force update even if the file hasn't changed
"""

import os
import sys
import urllib.request
import hashlib
import openpyxl
from datetime import datetime

# EDINET Taxonomy URL
EDINET_TAXONOMY_URL = "https://disclosure2dl.edinet-fsa.go.jp/guide/static/disclosure/download/ESE140115.xlsx"
TAXONOMY_FILE = "edinet_taxonomy_elements.xlsx"
OUTPUT_FILE = "edinet_taxonomy_dict.py"
HASH_FILE = ".edinet_taxonomy.hash"

def download_taxonomy():
    """Download EDINET taxonomy file"""
    print(f"Downloading EDINET taxonomy from: {EDINET_TAXONOMY_URL}")

    try:
        urllib.request.urlretrieve(EDINET_TAXONOMY_URL, TAXONOMY_FILE)
        print(f"✓ Downloaded: {TAXONOMY_FILE}")
        return True
    except Exception as e:
        print(f"✗ Download failed: {e}")
        return False

def calculate_file_hash(filename):
    """Calculate SHA256 hash of file"""
    if not os.path.exists(filename):
        return None

    sha256_hash = hashlib.sha256()
    with open(filename, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def check_if_update_needed(force=False):
    """Check if taxonomy file has changed"""
    if force:
        print("Force update requested")
        return True

    if not os.path.exists(TAXONOMY_FILE):
        print("Taxonomy file not found, update needed")
        return True

    current_hash = calculate_file_hash(TAXONOMY_FILE)

    if os.path.exists(HASH_FILE):
        with open(HASH_FILE, 'r') as f:
            saved_hash = f.read().strip()

        if current_hash == saved_hash:
            print("✓ Taxonomy file unchanged, no update needed")
            return False

    print("Taxonomy file has changed, update needed")
    return True

def save_hash():
    """Save current hash of taxonomy file"""
    current_hash = calculate_file_hash(TAXONOMY_FILE)
    with open(HASH_FILE, 'w') as f:
        f.write(current_hash)
    print(f"✓ Saved taxonomy hash: {current_hash[:16]}...")

def generate_dictionary():
    """Generate dictionary from EDINET taxonomy"""
    print(f"Generating dictionary from: {TAXONOMY_FILE}")

    try:
        wb = openpyxl.load_workbook(TAXONOMY_FILE, data_only=True)
        ws = wb['一般商工業']

        # Extract taxonomy dictionary
        edinet_dict = {}
        for row in ws.iter_rows(min_row=3, values_only=True):
            element_name = row[8]  # Column I: Element name
            namespace = row[7]     # Column H: Namespace
            jp_label = row[1]      # Column B: Japanese standard label

            if element_name and jp_label and namespace in ('jppfs_cor', 'jpigp_cor'):
                label = str(jp_label)
                # Shorten labels with "or loss" notation
                if '又は' in label and ('損失' in label or '損' in label) and '（△）' in label:
                    parts = label.split('又は')
                    if len(parts) == 2:
                        label = parts[0].strip()
                edinet_dict[element_name] = label

        print(f"✓ Extracted {len(edinet_dict)} items from EDINET taxonomy")

        # Custom mappings (IFRS variants, abbreviations, etc.)
        custom_mappings = {
            # Generic abbreviations
            'Notes': '注記',
            'Inventory': '棚卸資産',

            # Financial statement names
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

            # IFRS variants / Abbreviations
            'NetSalesIFRS': '売上収益',
            'RevenueIFRS': '売上収益',
            'CostOfSalesIFRS': '売上原価',
            'GrossProfitIFRS': '売上総利益',
            'SellingGeneralAndAdministrativeExpensesIFRS': '販売費及び一般管理費',
            'OtherOperatingIncome': 'その他の営業収益',
            'OtherOperatingExpenses': 'その他の営業費用',
            'OtherOperatingExpense': 'その他の営業費用',
            'OtherIncomeIFRS': 'その他の収益',
            'OtherExpensesIFRS': 'その他の費用',
            'OtherOperatingIncomeIFRS': 'その他の営業収益',
            'OtherOperatingExpensesIFRS': 'その他の営業費用',
            'ShareOfProfitLossOfInvestmentsAccountedForUsingEquityMethodIFRS': '持分法による投資利益',
            'OperatingProfitLossIFRS': '営業利益',
            'FinanceIncomeIFRS': '金融収益',
            'FinanceCostsIFRS': '金融費用',
            'ProfitLossBeforeTaxIFRS': '税引前当期利益',
            'IncomeTaxExpenseIFRS': '法人所得税費用',
            'ProfitLossIFRS': '当期利益',
            'ProfitLossAttributableToOwnersOfParentIFRS': '親会社の所有者に帰属する当期利益',
            'ProfitLossAttributableToNonControllingInterestsIFRS': '非支配持分',
            'BasicEarningsPerShareIFRS': '基本的１株当たり当期利益（円）',

            # J-GAAP variants / Abbreviations
            'OperatingProfit': '営業利益',
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
            'ProfitAttributableToOwnersOfParent': '親会社株主に帰属する当期純利益',
            'ProfitAttributableToNoncontrollingInterests': '非支配持分',
            'ProfitLossAttributableToNoncontrollingInterests': '非支配持分',
            'BasicEarningsPerShare': '基本的１株当たり当期純利益（円）',
            'BasicEarningsLossPerShare': '基本的１株当たり当期純利益（円）',
            'SellingGeneralAndAdministrativeExpense': '販売費及び一般管理費',
            'ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod': '持分法による投資利益',
        }

        # Merge dictionaries
        final_dict = {}
        final_dict.update(edinet_dict)
        final_dict.update(custom_mappings)

        print(f"✓ Total items: {len(final_dict)} (EDINET: {len(edinet_dict)}, Custom: {len(custom_mappings)})")

        return final_dict, len(edinet_dict), len(custom_mappings)

    except Exception as e:
        print(f"✗ Dictionary generation failed: {e}")
        import traceback
        traceback.print_exc()
        return None, 0, 0

def write_dictionary_file(final_dict, edinet_count, custom_count):
    """Write dictionary to Python file"""
    print(f"Writing dictionary to: {OUTPUT_FILE}")

    try:
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write('#!/usr/bin/env python3\n')
            f.write('# -*- coding: utf-8 -*-\n')
            f.write('"""\n')
            f.write('EDINET Taxonomy Dictionary\n')
            f.write('\n')
            f.write('Auto-generated from EDINET Official Taxonomy:\n')
            f.write(f'{EDINET_TAXONOMY_URL}\n')
            f.write('\n')
            f.write(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write(f'Total: {len(final_dict):,} items\n')
            f.write(f'- EDINET Official Taxonomy: {edinet_count:,} items\n')
            f.write(f'- Custom Mappings (IFRS variants, etc.): {custom_count} items\n')
            f.write('"""\n')
            f.write('\n')
            f.write('# EDINET Taxonomy common dictionary\n')
            f.write('# Maps element names to Japanese account labels\n')
            f.write('common_dict = {\n')

            # Sort and write entries
            for key in sorted(final_dict.keys()):
                value = final_dict[key]
                # Escape single quotes
                value_escaped = value.replace("'", "\\'")
                f.write(f"    '{key}': '{value_escaped}',\n")

            f.write('}\n')

        print(f"✓ Dictionary written to: {OUTPUT_FILE}")
        return True

    except Exception as e:
        print(f"✗ File write failed: {e}")
        return False

def main():
    """Main function"""
    print("=" * 80)
    print("EDINET Taxonomy Dictionary Auto-Update")
    print("=" * 80)
    print()

    force = '--force' in sys.argv

    # Step 1: Download taxonomy file
    if not os.path.exists(TAXONOMY_FILE) or force:
        if not download_taxonomy():
            print("\n✗ Update failed: Could not download taxonomy file")
            return 1

    # Step 2: Check if update is needed
    if not check_if_update_needed(force):
        print("\n✓ No update needed")
        return 0

    # Step 3: Generate dictionary
    final_dict, edinet_count, custom_count = generate_dictionary()
    if final_dict is None:
        print("\n✗ Update failed: Could not generate dictionary")
        return 1

    # Step 4: Write dictionary file
    if not write_dictionary_file(final_dict, edinet_count, custom_count):
        print("\n✗ Update failed: Could not write dictionary file")
        return 1

    # Step 5: Save hash
    save_hash()

    print()
    print("=" * 80)
    print("✓ Dictionary update completed successfully!")
    print("=" * 80)

    return 0

if __name__ == '__main__':
    sys.exit(main())
